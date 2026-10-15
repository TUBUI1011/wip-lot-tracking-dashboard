"use client";
// import React from "react";
import { ChangeEvent, useMemo, useState } from "react";
import * as XLSX from "xlsx";

// AREA: WIP LOT TRACKING DASHBOARD
// Khai báo các kiểu dữ liệu và hằng số cho bảng WIP LOT TRACKING DASHBOARD
// WipLot là kiểu dữ liệu đại diện cho một lô trong bảng WIP LOT TRACKING DASHBOARD, 
// bao gồm các thông tin như số lô, nhóm sản phẩm, thao tác, số lượng, chủ sở hữu, cờ giữ, cờ có thể giao, nền tảng, vị trí, số ngày trong thao tác, 
// các chỉ số MET và trạng thái cửa sổ SIT MET.
type WipLot = {
  lot: string; productGroup3: string; operation: string; operationDesc: string;
  qty: number; owner: string; holdFlag: string; shippableFlag: string;
  platform: string; location: string; daysInOperation: number | null;
  met1: number | null; met2: number | null; met3: number | null;
  sitMetWindow: string; sitMetTimeRemaining: number | null;
  sitMetWindowStatus: string; metDrbTimeRemaining: number | null;
};

// RawRow là kiểu dữ liệu đại diện cho một hàng dữ liệu thô trong bảng Excel,
//  được biểu diễn dưới dạng một đối tượng với các khóa là tên cột và giá trị là dữ liệu tương ứng.
type RawRow = Record<string, unknown>;
const REQUIRED_COLUMNS = ["Lot", "Operation"]; // Các cột bắt buộc trong bảng Excel
const HEADER_MAP: Record<keyof WipLot, string[]> = { // Bản đồ ánh xạ giữa các trường dữ liệu WipLot và các tiêu đề cột tương ứng trong bảng Excel
  lot: ["Lot", "Lot Number", "Lot ID"],
  productGroup3: ["Product Group 3", "Product Group3"],
  operation: ["Operation", "Operation Number", "Oper"],
  operationDesc: ["Operation Desc", "Operation Description", "Oper Desc"],
  qty: ["Qty", "Quantity"], owner: ["Owner"], holdFlag: ["Hold Flag"],
  shippableFlag: ["Shippable Flag"], platform: ["Platform"], location: ["Location"],
  daysInOperation: ["Days In Operation"], met1: ["MET1"], met2: ["MET2"], met3: ["MET3"],
  sitMetWindow: ["SIT MET Window"], sitMetTimeRemaining: ["SIT MET Time Remaining"],
  sitMetWindowStatus: ["SIT MET Window Status"], metDrbTimeRemaining: ["MET DRB Time Remaining"],
};

// KHU VỰC CÁC HÀM
// Các hàm tiện ích để xử lý dữ liệu WipLot và RawRow, bao gồm chuẩn hóa tiêu đề, tìm tiêu đề, 
// lấy giá trị văn bản, giá trị số, định dạng số, và xác định lớp CSS cho cờ và trạng thái SIT.

// Hàm NormalizeHeader: chuẩn hóa tiêu đề cột bằng cách loại bỏ khoảng trắng thừa, 
// chuyển sang chữ thường và thay thế các ký tự gạch dưới hoặc gạch ngang bằng khoảng trắng.
function normalizeHeader(value: unknown) {
  return String(value ?? "").trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ");
}
// Hàm FindHeader: tìm tiêu đề cột phù hợp trong danh sách tiêu đề dựa trên các ứng cử viên.
function findHeader(headers: string[], candidates: string[]) {
  const normalized = candidates.map(normalizeHeader);
  return headers.find((header) => normalized.includes(normalizeHeader(header)));
}
// Hàm TextValue: lấy giá trị văn bản từ một giá trị bất kỳ, trả về chuỗi rỗng nếu giá trị là null hoặc undefined.
function textValue(value: unknown) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}
// Hàm NumberValue: lấy giá trị số từ một giá trị bất kỳ, trả về null nếu giá trị không hợp lệ.
// dùng để lấy giá trị số từ một giá trị bất kỳ, trả về null nếu giá trị không hợp lệ.
function numberValue(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const parsed = Number(String(value).replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : null;
}
// Hàm fmt: định dạng giá trị số với số chữ số thập phân cố định, trả về "—" nếu giá trị không hợp lệ.
function fmt(value: number | null | undefined, decimals = 2) {
  if (value === null || value === undefined || !Number.isFinite(value)) return "—";
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: decimals }).format(value);
}
// Hàm flagClass: xác định lớp CSS cho cờ dựa trên giá trị chuỗi ("y", "yes", "true" => flag-yes; "n", "no", "false" => flag-no; khác => flag).
function flagClass(value: string) {
  const v = value.toLowerCase();
  if (["y", "yes", "true"].includes(v)) return "flag flag-yes";
  if (["n", "no", "false"].includes(v)) return "flag flag-no";
  return "flag";
}
// Hàm sitClass: xác định lớp CSS cho trạng thái SIT dựa trên trạng thái cửa sổ MET và thời gian còn lại.
// Trả về "sit-badge sit-danger" nếu trạng thái MET đã hết hạn hoặc thời gian còn lại âm.
// Trả về "sit-badge sit-warning" nếu thời gian còn lại âm nhưng trạng thái MET chưa hết hạn.
// Trả về "sit-badge sit-normal" nếu trạng thái MET hợp lệ.
// Trả về "sit-badge" nếu không có thông tin trạng thái MET.
function sitClass(row: WipLot) {
  const status = row.sitMetWindowStatus.toLowerCase();
  if (status.includes("over") || status.includes("expired") || (row.sitMetTimeRemaining !== null && row.sitMetTimeRemaining < 0)) return "sit-badge sit-danger";
  if (row.sitMetTimeRemaining !== null && row.sitMetTimeRemaining < 0) return "sit-badge sit-warning";
  if (row.sitMetWindowStatus) return "sit-badge sit-normal";
  return "sit-badge";
}
// Hàm parseWorkbook: phân tích tệp Excel và ánh xạ dữ liệu vào mảng WipLot.
async function parseWorkbook(file: File): Promise<WipLot[]> {
  const bytes = await file.arrayBuffer();
  const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("The workbook does not contain a worksheet.");
  const rows = XLSX.utils.sheet_to_json<RawRow>(sheet, { defval: "" });
  if (!rows.length) return [];
  const headers = Object.keys(rows[0]);
  const lotHeader = findHeader(headers, HEADER_MAP.lot);
  const operationHeader = findHeader(headers, HEADER_MAP.operation);
  if (!lotHeader || !operationHeader) {
    throw new Error(`Could not find the required columns: ${REQUIRED_COLUMNS.join(" and ")}. Please check the first worksheet header row.`);
  }
  const headerFor = (key: keyof WipLot) => findHeader(headers, HEADER_MAP[key]);
  const mapped = rows.map((row): WipLot => {
    const get = (key: keyof WipLot) => {
      const header = headerFor(key);
      return header ? row[header] : "";
    };
    return {
      lot: textValue(row[lotHeader]), productGroup3: textValue(get("productGroup3")),
      operation: textValue(row[operationHeader]), operationDesc: textValue(get("operationDesc")),
      qty: numberValue(get("qty")) ?? 0, owner: textValue(get("owner")), holdFlag: textValue(get("holdFlag")),
      shippableFlag: textValue(get("shippableFlag")), platform: textValue(get("platform")), location: textValue(get("location")),
      daysInOperation: numberValue(get("daysInOperation")), met1: numberValue(get("met1")), met2: numberValue(get("met2")), met3: numberValue(get("met3")),
      sitMetWindow: textValue(get("sitMetWindow")), sitMetTimeRemaining: numberValue(get("sitMetTimeRemaining")),
      sitMetWindowStatus: textValue(get("sitMetWindowStatus")), metDrbTimeRemaining: numberValue(get("metDrbTimeRemaining")),
    };
  }).filter((row) => /^[XU]/i.test(row.lot.trim()));
  return mapped;
}

export default function Home() {
  const [data, setData] = useState<WipLot[]>([]);
  const [fileName, setFileName] = useState("");
  const [notice, setNotice] = useState("Upload an ION WIP Lot List Excel file to begin. The file is parsed in this browser session.");
  const [isError, setIsError] = useState(false);
  const [search, setSearch] = useState("");
  const [selectedOperations, setSelectedOperations] = useState<string[]>([]);
  const [operationSearch, setOperationSearch] = useState("");
  const [filters, setFilters] = useState({ productGroup3: "", owner: "", platform: "", location: "", hold: "all", shippable: "all", sitStatus: "all" });
  const [sortBy, setSortBy] = useState<keyof WipLot>("sitMetTimeRemaining");
  const [sortDirection, setSortDirection] = useState<"asc" | "desc">("asc");
  const [page, setPage] = useState(1);
  const [busy, setBusy] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [activeView, setActiveView] = useState<"dashboard" | "overall">("dashboard");
  const [overallPlatform, setOverallPlatform] = useState("all");
  const [overallSearch, setOverallSearch] = useState("");
  const pageSize = 25;

  const operations = useMemo(() => [...new Set(data.map((row) => row.operation).filter(Boolean))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })), [data]);
  const distinct = useMemo(() => ({
    productGroup3: [...new Set(data.map((r) => r.productGroup3).filter(Boolean))].sort(),
    owner: [...new Set(data.map((r) => r.owner).filter(Boolean))].sort(),
    platform: [...new Set(data.map((r) => r.platform).filter(Boolean))].sort(),
    location: [...new Set(data.map((r) => r.location).filter(Boolean))].sort(),
  }), [data]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return data.filter((row) => {
      const searchMatch = !q || [row.lot, row.productGroup3, row.operation, row.operationDesc, row.owner, row.platform, row.location, row.sitMetWindowStatus, row.sitMetWindow].some((v) => v.toLowerCase().includes(q));
      const operationMatch = selectedOperations.length === 0 || selectedOperations.includes(row.operation);
      const sitMatch = filters.sitStatus === "all" || (filters.sitStatus === "missing" ? !row.sitMetWindowStatus : filters.sitStatus === "risk" ? (row.sitMetTimeRemaining !== null && row.sitMetTimeRemaining <= 8) || row.sitMetWindowStatus.toLowerCase().includes("over") : row.sitMetWindowStatus.toLowerCase().includes(filters.sitStatus));
      return searchMatch && operationMatch && sitMatch &&
        (!filters.productGroup3 || row.productGroup3 === filters.productGroup3) &&
        (!filters.owner || row.owner === filters.owner) && (!filters.platform || row.platform === filters.platform) &&
        (!filters.location || row.location === filters.location) &&
        (filters.hold === "all" || row.holdFlag.toLowerCase() === filters.hold) &&
        (filters.shippable === "all" || row.shippableFlag.toLowerCase() === filters.shippable);
    });
  }, [data, search, selectedOperations, filters]);

  const sorted = useMemo(() => [...filtered].sort((a, b) => {
    const av = a[sortBy]; const bv = b[sortBy];
    let result = 0;
    if (typeof av === "number" || typeof bv === "number") result = (typeof av === "number" ? av : Number.POSITIVE_INFINITY) - (typeof bv === "number" ? bv : Number.POSITIVE_INFINITY);
    else result = String(av ?? "").localeCompare(String(bv ?? ""), undefined, { numeric: true, sensitivity: "base" });
    return sortDirection === "asc" ? result : -result;
  }), [filtered, sortBy, sortDirection]);

  const totalPages = Math.max(1, Math.ceil(sorted.length / pageSize));
  const visibleRows = sorted.slice((page - 1) * pageSize, page * pageSize);
  const holdLots = filtered.filter((row) => ["y", "yes", "true"].includes(row.holdFlag.toLowerCase())).length;
  const shippableLots = filtered.filter((row) => ["y", "yes", "true"].includes(row.shippableFlag.toLowerCase())).length;
  const sitRiskLots = filtered.filter((row) => (row.sitMetTimeRemaining !== null && row.sitMetTimeRemaining < 0) || row.sitMetWindowStatus.toLowerCase().includes("over") || row.sitMetWindowStatus.toLowerCase().includes("expired")).length;
  const totalQty = filtered.reduce((sum, row) => sum + row.qty, 0);
  const filteredOperations = useMemo(() => operations.filter((op) => op.toLowerCase().includes(operationSearch.toLowerCase())), [operations, operationSearch]);
  const overallWip = useMemo(() => {
    const groups = new Map<string, { platform: string; product: string; qty: number; lots: number }>();
    data.forEach((row) => {
      const platform = row.platform || "Unspecified platform";
      const product = row.productGroup3 || "Unspecified product";
      const key = `${platform}\u0000${product}`;
      const current = groups.get(key) || { platform, product, qty: 0, lots: 0 };
      current.qty += row.qty;
      current.lots += 1;
      groups.set(key, current);
    });
    const query = overallSearch.trim().toLowerCase();
    return [...groups.values()]
      .filter((item) => (overallPlatform === "all" || item.platform === overallPlatform) &&
        (!query || item.platform.toLowerCase().includes(query) || item.product.toLowerCase().includes(query)))
      .sort((a, b) => b.qty - a.qty || a.platform.localeCompare(b.platform) || a.product.localeCompare(b.product));
  }, [data, overallPlatform, overallSearch]);
  const overallTotalQty = useMemo(() => data.reduce((sum, row) => sum + row.qty, 0), [data]);
  const overallTotalLots = data.length;
  const overallPlatforms = useMemo(() => [...new Set(data.map((row) => row.platform).filter(Boolean))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })), [data]);

  async function handleUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    setBusy(true); setIsError(false);
    try {
      const parsed = await parseWorkbook(file);
      if (!parsed.length) throw new Error("No data rows were found in the first worksheet.");
      setData(parsed); setFileName(file.name); setSelectedOperations([]); setPage(1);
      setNotice(`Loaded ${fmt(parsed.length, 0)} valid WIP lots from ${file.name} (Lot starts with X or U). Choose one or more operations to focus the view.`);
    } catch (error) {
      setIsError(true); setNotice(error instanceof Error ? error.message : "Unable to read this Excel file.");
    } finally {
      setBusy(false); event.target.value = "";
    }
  }
  function updateFilter(key: keyof typeof filters, value: string) { setFilters((old) => ({ ...old, [key]: value })); setPage(1); }
  function toggleOperation(operation: string) {
    setSelectedOperations((old) => old.includes(operation) ? old.filter((item) => item !== operation) : [...old, operation]);
    setPage(1);
  }
  function sortColumn(key: keyof WipLot) {
    if (sortBy === key) setSortDirection((old) => old === "asc" ? "desc" : "asc");
    else { setSortBy(key); setSortDirection("asc"); }
  }
  function resetFilters() {
    setSearch(""); setSelectedOperations([]); setOperationSearch("");
    setFilters({ productGroup3: "", owner: "", platform: "", location: "", hold: "all", shippable: "all", sitStatus: "all" }); setPage(1);
  }

  // HÀM exportCsv: xuất dữ liệu đã lọc và sắp xếp ra tệp CSV.
  function exportCsv() {
    const columns: (keyof WipLot)[] = ["lot", "productGroup3", "operation", "operationDesc", "qty", "owner", "holdFlag", "shippableFlag", "platform", "location", "daysInOperation", "met1", "met2", "met3", "sitMetWindow", "sitMetTimeRemaining", "sitMetWindowStatus", "metDrbTimeRemaining"];
    const headerLabels = ["Lot", "Product Group 3", "Operation", "Operation Desc", "Qty", "Owner", "Hold Flag", "Shippable Flag", "Platform", "Location", "Days In Operation", "MET1", "MET2", "MET3", "SIT MET Window", "SIT MET Time Remaining", "SIT MET Window Status", "MET DRB Time Remaining"];
    const escape = (value: unknown) => `"${String(value ?? "").replace(/"/g, '""')}"`;
    const csv = [headerLabels.map(escape).join(","), ...sorted.map((row) => columns.map((key) => escape(row[key])).join(","))].join("\r\n");
    const blob = new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob); const link = document.createElement("a");
    link.href = url; link.download = "wip-lot-filtered.csv"; link.click(); URL.revokeObjectURL(url);
  }
  // HÀM th: tạo tiêu đề cột bảng với khả năng sắp xếp.
  const th = (label: string, key: keyof WipLot, numeric = false) => <th key={key} className={numeric ? "numeric" : ""}><button className="sort-button" onClick={() => sortColumn(key)}>{label}{sortBy === key ? (sortDirection === "asc" ? " ↑" : " ↓") : " ↕"}</button></th>;

  return <main className={`app-shell ${sidebarOpen ? "sidebar-open" : "sidebar-collapsed"}`}>
    <aside className="sidebar">
      <div className="brand"><div className="brand-mark">W</div><div className="brand-copy"><strong>WIP EXECUTE</strong><span>LOT TRACKING</span></div></div>
      <div className="sidebar-label">WORKSPACE</div>
      <nav className="side-nav">
        <button className={`nav-item ${activeView === "dashboard" ? "active" : ""}`} onClick={() => setActiveView("dashboard")} title="WIP Dashboard"><span>▦</span><b>WIP Dashboard</b></button>
        <button className={`nav-item ${activeView === "overall" ? "active" : ""}`} onClick={() => setActiveView("overall")} title="Overall WIP"><span>▤</span><b>Overall WIP</b></button>
        <button className="nav-item" onClick={() => { setActiveView("dashboard"); document.getElementById("wip-lot-table")?.scrollIntoView({ behavior: "smooth", block: "start" }); }} title="Lot Explorer"><span>⌕</span><b>Lot Explorer</b></button>
        <button className="nav-item" onClick={() => { setActiveView("dashboard"); document.getElementById("sit-met-filters")?.scrollIntoView({ behavior: "smooth", block: "center" }); }} title="SIT MET Watch"><span>◷</span><b>SIT MET Watch</b></button>
      </nav>
      <div className="sidebar-bottom"><span className="live-dot" /> LOCAL SESSION <p>Excel data stays in this browser session. No upload to a database.</p><small>PHASE 1 · WIP EXECUTE</small></div>
    </aside>
    <section className="main-panel">
      <header className="page-header"><div className="header-title"><button className="sidebar-toggle" onClick={() => setSidebarOpen((open) => !open)} aria-label={sidebarOpen ? "Collapse sidebar" : "Expand sidebar"} title={sidebarOpen ? "Collapse toolbar" : "Open toolbar"}>{sidebarOpen ? "‹" : "›"}</button><div><div className="eyebrow">MANUFACTURING · EXECUTION VIEW</div><h1>{activeView === "overall" ? "Overall WIP" : "WIP Lot Tracking Dashboard"}</h1><p>{activeView === "overall" ? "Total WIP quantity by Platform and Product, ranked from highest to lowest." : "Find the lots that need attention without digging through multiple pivots."}</p></div></div><div className="session-pill"><span /> BROWSER-LOCAL PROCESSING</div></header>

      <section className="upload-card"><div className="excel-icon">X</div><div className="upload-copy"><strong>Import ION WIP Lot List</strong><p>{fileName || "Choose the latest Excel export to refresh the dashboard."}</p><small>Supports .xlsx and .xls · Uses the first worksheet · New import replaces current data</small></div><label className={`primary-button ${busy ? "disabled" : ""}`}>{busy ? "Reading Excel…" : "＋ Import Excel"}<input type="file" accept=".xlsx,.xls,.xlsm" onChange={handleUpload} disabled={busy} /></label></section>
      <div className={`notice ${isError ? "notice-error" : ""}`} role="status">{notice}</div>

      {activeView === "overall" ? <section className="overall-view">
        <div className="overall-summary-grid">
          <article className="overall-summary"><span>OVERALL WIP QTY</span><strong>{fmt(overallTotalQty, 0)}</strong><small>Sum of Qty across all imported rows</small></article>
          <article className="overall-summary"><span>WIP LOTS</span><strong>{fmt(overallTotalLots, 0)}</strong><small>Rows in the imported WIP list</small></article>
          <article className="overall-summary"><span>PLATFORMS</span><strong>{fmt(overallPlatforms.length, 0)}</strong><small>Distinct platform values</small></article>
          <article className="overall-summary"><span>PRODUCT COMBINATIONS</span><strong>{fmt(overallWip.length, 0)}</strong><small>Platform and Product groups shown</small></article>
        </div>
        <section className="panel overall-panel">
          <div className="table-heading"><div><h2>Overall WIP by Platform &amp; Product</h2><p>Sorted by total Qty from highest to lowest. Values are calculated from the full imported file, not the dashboard filters.</p></div><span className="table-tag">DESCENDING QTY</span></div>
          <div className="overall-controls"><label className="search-field"><span>⌕</span><input value={overallSearch} onChange={(e) => setOverallSearch(e.target.value)} placeholder="Search platform or product…" /></label><label className="select-field"><span>Platform</span><select value={overallPlatform} onChange={(e) => setOverallPlatform(e.target.value)}><option value="all">All platforms</option>{overallPlatforms.map((platform) => <option key={platform} value={platform}>{platform}</option>)}</select></label><button className="text-button" onClick={() => { setOverallPlatform("all"); setOverallSearch(""); }}>Reset</button></div>
          <div className="table-scroll overall-table-scroll"><table><thead><tr><th>Rank</th><th>Platform</th><th>Product Group 3</th><th className="numeric">Total Qty</th><th className="numeric">Lot Count</th><th>Share of Overall Qty</th></tr></thead><tbody>{overallWip.map((item, index) => <tr key={`${item.platform}-${item.product}`}><td><span className={`rank-badge ${index < 3 ? "top-rank" : ""}`}>{index + 1}</span></td><td className="platform-cell">{item.platform}</td><td>{item.product}</td><td className="numeric overall-qty">{fmt(item.qty, 0)}</td><td className="numeric">{fmt(item.lots, 0)}</td><td><div className="share-cell"><div className="share-track"><span style={{ width: `${overallTotalQty > 0 ? Math.min(100, item.qty / overallTotalQty * 100) : 0}%` }} /></div><small>{overallTotalQty > 0 ? `${(item.qty / overallTotalQty * 100).toFixed(1)}%` : "0.0%"}</small></div></td></tr>)}{!overallWip.length && <tr><td colSpan={6} className="no-results">Import an Excel file or adjust the Overall WIP filters.</td></tr>}</tbody></table></div>
          <div className="table-footer"><span>Showing {fmt(overallWip.length, 0)} Platform / Product group(s)</span><span>Overall Qty: <strong>{fmt(overallTotalQty, 0)}</strong></span></div>
        </section>
      </section> : <>
      <section className="kpi-grid">
        <article className="kpi-card"><div className="kpi-icon blue">▦</div><div><span>WIP LOTS</span><strong>{fmt(filtered.length, 0)}</strong><small>{selectedOperations.length ? `${selectedOperations.length} operation(s) selected` : "All operations"}</small></div></article>
        <article className="kpi-card"><div className="kpi-icon purple">∑</div><div><span>TOTAL QTY</span><strong>{fmt(totalQty, 0)}</strong><small>Sum of Qty in current view</small></div></article>
        <article className="kpi-card"><div className="kpi-icon orange">!</div><div><span>ON HOLD</span><strong>{fmt(holdLots, 0)}</strong><small>Lots with Hold Flag = Y</small></div></article>
        <article className="kpi-card"><div className="kpi-icon green">◷</div><div><span>SIT MET OVERDUE</span><strong>{fmt(sitRiskLots, 0)}</strong><small>{fmt(shippableLots, 0)} shippable lot(s) in view</small></div></article>
      </section>

      <section className="operation-panel panel"><div className="panel-heading"><div><h2>Focus Operations</h2><p>Select one or more operations owned by your module. Leave empty to view all.</p></div><div className="operation-actions"><span>{selectedOperations.length} selected</span><button className="text-button" onClick={() => { setSelectedOperations([]); setPage(1); }}>Clear</button></div></div>
        <input className="operation-search" value={operationSearch} onChange={(e) => setOperationSearch(e.target.value)} placeholder="Search operation number or description…" />
        {data.length ? <div className="operation-list">{filteredOperations.slice(0, 80).map((operation) => { const sample = data.find((r) => r.operation === operation); const count = data.filter((r) => r.operation === operation).length; return <label key={operation} className={`operation-option ${selectedOperations.includes(operation) ? "selected" : ""}`}><input type="checkbox" checked={selectedOperations.includes(operation)} onChange={() => toggleOperation(operation)} /><span className="op-number">{operation}</span><span className="op-description">{sample?.operationDesc || "No description"}</span><span className="op-count">{fmt(count, 0)}</span></label>; })}{filteredOperations.length > 80 && <div className="operation-more">Showing first 80 of {fmt(filteredOperations.length, 0)} operations. Search to narrow the list.</div>}</div> : <div className="empty-inline">Import an Excel file to load available operations.</div>}
      </section>

      <section className="filters-panel panel" id="sit-met-filters"><div className="panel-heading"><div><h2>Filters</h2><p>Search across lot, operation, owner, platform, location and SIT MET fields.</p></div><button className="text-button" onClick={resetFilters}>Reset all</button></div>
        <div className="filter-grid"><label className="search-field"><span>⌕</span><input value={search} onChange={(e) => { setSearch(e.target.value); setPage(1); }} placeholder="Search lot, operation, owner…" /></label>
          <label className="select-field"><span>Product Group 3</span><select value={filters.productGroup3} onChange={(e) => updateFilter("productGroup3", e.target.value)}><option value="">All groups</option>{distinct.productGroup3.map((v) => <option key={v}>{v}</option>)}</select></label>
          <label className="select-field"><span>Owner</span><select value={filters.owner} onChange={(e) => updateFilter("owner", e.target.value)}><option value="">All owners</option>{distinct.owner.map((v) => <option key={v}>{v}</option>)}</select></label>
          <label className="select-field"><span>Platform</span><select value={filters.platform} onChange={(e) => updateFilter("platform", e.target.value)}><option value="">All platforms</option>{distinct.platform.map((v) => <option key={v}>{v}</option>)}</select></label>
          <label className="select-field"><span>Location</span><select value={filters.location} onChange={(e) => updateFilter("location", e.target.value)}><option value="">All locations</option>{distinct.location.map((v) => <option key={v}>{v}</option>)}</select></label>
          <label className="select-field"><span>Hold Flag</span><select value={filters.hold} onChange={(e) => updateFilter("hold", e.target.value)}><option value="all">All</option><option value="y">Hold = Y</option><option value="n">Hold = N</option></select></label>
          <label className="select-field"><span>Shippable Flag</span><select value={filters.shippable} onChange={(e) => updateFilter("shippable", e.target.value)}><option value="all">All</option><option value="y">Shippable = Y</option><option value="n">Shippable = N</option></select></label>
          <label className="select-field"><span>SIT MET status</span><select value={filters.sitStatus} onChange={(e) => updateFilter("sitStatus", e.target.value)}><option value="all">All statuses</option><option value="risk">Overdue / negative time</option><option value="missing">Missing status</option><option value="mvin">MVIN</option></select></label>
        </div>
      </section>

      <section className="table-panel panel" id="wip-lot-table"><div className="table-heading"><div><h2>WIP Lot List</h2><p>{fmt(sorted.length, 0)} lot(s) match the current filters · sorted by {String(sortBy)}</p></div><div className="table-tools"><span className="table-tag">ION DATA</span><button className="export-button" onClick={exportCsv} disabled={!sorted.length}>↓ Export CSV</button></div></div>
        <div className="table-scroll"><table><thead><tr>{th("Lot", "lot")}{th("Product Group 3", "productGroup3")}{th("Operation", "operation")}{th("Operation Desc", "operationDesc")}{th("Qty", "qty", true)}{th("Owner", "owner")}{th("Hold", "holdFlag")}{th("Shippable", "shippableFlag")}{th("Platform", "platform")}{th("Location", "location")}{th("Days in Op.", "daysInOperation", true)}{th("SIT MET Window", "sitMetWindow")}{th("SIT MET Remaining", "sitMetTimeRemaining", true)}{th("SIT MET Status", "sitMetWindowStatus")}{th("MET DRB Remaining", "metDrbTimeRemaining", true)}</tr></thead>
          <tbody>{visibleRows.map((row, index) => <tr key={`${row.lot}-${row.operation}-${index}`}><td className="lot-cell">{row.lot || "—"}</td><td>{row.productGroup3 || "—"}</td><td><button className="operation-link" onClick={() => toggleOperation(row.operation)} title="Click to toggle this operation">{row.operation || "—"}</button></td><td className="desc-cell">{row.operationDesc || "—"}</td><td className="numeric qty-cell">{fmt(row.qty, 0)}</td><td>{row.owner || "—"}</td><td><span className={flagClass(row.holdFlag)}>{row.holdFlag || "—"}</span></td><td><span className={flagClass(row.shippableFlag)}>{row.shippableFlag || "—"}</span></td><td>{row.platform || "—"}</td><td className="location-cell">{row.location || "—"}</td><td className="numeric">{fmt(row.daysInOperation, 2)}</td><td>{row.sitMetWindow || "—"}</td><td className="numeric"><strong className={row.sitMetTimeRemaining !== null && row.sitMetTimeRemaining < 0 ? "time-risk" : ""}>{fmt(row.sitMetTimeRemaining)}</strong></td><td><span className={sitClass(row)}>{row.sitMetWindowStatus || "No data"}</span></td><td className="numeric">{fmt(row.metDrbTimeRemaining)}</td></tr>)}
          {!visibleRows.length && <tr><td colSpan={15} className="no-results">{data.length ? "No lots match these filters. Try resetting filters or selecting different operations." : "Import the ION Excel file to see WIP lots here."}</td></tr>}</tbody></table></div>
        <div className="table-footer"><span>Showing {sorted.length ? (page - 1) * pageSize + 1 : 0}–{Math.min(page * pageSize, sorted.length)} of {fmt(sorted.length, 0)} lots</span><div className="pagination"><button disabled={page <= 1} onClick={() => setPage(1)} title="First page">«</button><button disabled={page <= 1} onClick={() => setPage((p) => p - 1)} title="Previous page">‹</button><span>{page} / {totalPages}</span><button disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)} title="Next page">›</button><button disabled={page >= totalPages} onClick={() => setPage(totalPages)} title="Last page">»</button></div></div>
      </section>
      </>}
      <footer className="footer"><span>WIP EXECUTE · PHASE 1</span><span>Browser-side Excel parsing · No database · No automatic upload</span></footer>
    </section>
  </main>;
}
