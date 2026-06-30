import React, { useMemo, useState } from 'react';
import { AlertCircle, CheckCircle2, Download, FileSpreadsheet, PlayCircle, Search } from 'lucide-react';

const API_BASE = import.meta.env.VITE_API_BASE || ((window.location.port === '5173' || window.location.port === '5174') ? 'http://127.0.0.1:8000' : window.location.origin);
const TAB_ORDER = ['Summary', '119 vs Mgmt C', 'Mgmt C vs 119', 'All Exceptions', 'GS Discount', 'Issues', 'Approval'];

function UploadCard({ title, file, onChange }) {
  return (
    <label className="upload-card">
      <div className="upload-icon"><FileSpreadsheet size={30} /></div>
      <div>
        <h3>{title}</h3>
        <p>{file ? file.name : 'Choose an Excel workbook (.xlsx, .xlsm, .xls)'}</p>
      </div>
      <input type="file" accept=".xlsx,.xlsm,.xls" onChange={e => onChange(e.target.files?.[0] || null)} />
    </label>
  );
}

function MetricCards({ metrics }) {
  const keys = [
    'Matched trades with no exception',
    'Brokerage differences found on 119 vs Mgmt C',
    'Brokerage differences found on Mgmt C vs 119',
    'Trades in 119 source report missing from Mgmt C source report',
    'Trades in Mgmt C source report missing from 119 source report',
    'Final Issues rows after Fees filter and Discount % = 0% rule'
  ];
  return (
    <div className="metric-grid">
      {keys.map(k => <div className="metric-card" key={k}><span>{k}</span><strong>{metrics?.[k] ?? 0}</strong></div>)}
    </div>
  );
}

function TableView({ rows, search, showInternal }) {
  const filtered = useMemo(() => {
    if (!rows?.length) return [];
    if (!search) return rows;
    const s = search.toLowerCase();
    return rows.filter(row => Object.values(row).some(v => String(v ?? '').toLowerCase().includes(s)));
  }, [rows, search]);
  if (!rows?.length) return <div className="empty">No rows for this tab.</div>;
  let headers = Object.keys(rows[0]);
  if (!showInternal) headers = headers.filter(h => !['119 Row', 'Mgmt C Row', 'Sum of Fees', 'GS Involvement Flag'].includes(h));
  return (
    <div className="table-wrap">
      <table>
        <thead><tr>{headers.map(h => <th key={h}>{h}</th>)}</tr></thead>
        <tbody>
          {filtered.map((row, i) => <tr key={i}>{headers.map(h => <td key={h}>{formatValue(h, row[h])}</td>)}</tr>)}
        </tbody>
      </table>
      <div className="table-foot">Showing {filtered.length} of {rows.length} rows. API preview is capped for screen speed; Excel export includes the workbook output.</div>
    </div>
  );
}

function formatValue(header, value) {
  if (value === null || value === undefined || value === '') return '';
  if (header.includes('%') && typeof value === 'number') return `${(value * 100).toFixed(2)}%`;
  if (['119 Brokerage','Mgmt C Brokerage','Difference 119 - Mgmt C','Sum of Fees','Fees'].includes(header) && typeof value === 'number') return value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return String(value);
}

function SummaryTab({ result }) {
  return (
    <div className="summary-panel">
      <h2>119 vs Mgmt C Month End Reconciliation Summary</h2>
      <p>This page explains the source files, matching rules, exception counts, and final output tabs.</p>
      <div className="summary-grid">
        {Object.entries(result.summary || {}).map(([k, v]) => <div key={k}><b>{k}</b><span>{String(v)}</span></div>)}
      </div>
      <h3>Formula definitions</h3>
      <ul>
        <li>Difference 119 - Mgmt C = 119 Brokerage - Mgmt C Brokerage</li>
        <li>Discount % = IF(Sum of Discount = 0 or Sum of Net in LC = 0, 0, Sum of Discount / Sum of Net in LC)</li>
        <li>Fees = IF(ABS(Difference 119 - Mgmt C) &gt; 0.1, Difference 119 - Mgmt C - Sum of Fees, 0)</li>
        <li>GS Involvement Flag = populated when absolute Discount % is between 3% and 10% inclusive</li>
      </ul>
      <h3>Metrics</h3>
      <div className="metric-table">
        {Object.entries(result.metrics || {}).map(([k, v]) => <div key={k}><span>{k}</span><strong>{String(v)}</strong></div>)}
      </div>
    </div>
  );
}

function ApprovalTab() {
  return (
    <div className="approval-shell">
      <div className="approval-hero">
        <div>
          <span className="approval-eyebrow">Month-End Control Evidence</span>
          <h2>Reconciliation Sign-Off</h2>
          <p>This reconciliation has been reviewed and approved by:</p>
        </div>
        <div className="approval-stamp">Ready for Review</div>
      </div>

      <div className="approval-form-grid">
        <label>
          <span>Approver Name</span>
          <input placeholder="Enter approver name" />
        </label>
        <label>
          <span>Date Signed</span>
          <input type="date" />
        </label>
      </div>

      <label className="approval-comments">
        <span>Comments</span>
        <textarea rows="7" placeholder="Add approval notes, open items, or review comments." />
      </label>

      <div className="approval-checklist">
        <h3>Reviewer Checklist</h3>
        <div><span></span>Summary metrics reviewed</div>
        <div><span></span>Issues tab reviewed and comments completed</div>
        <div><span></span>GS Discount tab reviewed</div>
        <div><span></span>Final workbook approved for month-end evidence</div>
      </div>
    </div>
  );
}

export default function App() {
  const [raw119, setRaw119] = useState(null);
  const [mgmtc, setMgmtc] = useState(null);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState(null);
  const [activeTab, setActiveTab] = useState('Summary');
  const [search, setSearch] = useState('');
  const [showInternal, setShowInternal] = useState(false);

  async function runRecon() {
    setError(''); setRunning(true); setResult(null);
    const form = new FormData();
    form.append('raw119', raw119);
    form.append('mgmtc', mgmtc);
    try {
      const res = await fetch(`${API_BASE}/api/reconcile`, { method: 'POST', body: form });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Reconciliation failed.');
      setResult(data); setActiveTab('Summary');
    } catch (e) {
      setError(e.message);
    } finally {
      setRunning(false);
    }
  }

  function download() {
    if (!result) return;
    window.location.href = `${API_BASE}/api/download/${result.resultId}`;
  }

  return (
    <div>
      <nav><div className="brand"><FileSpreadsheet />119 vs Mgmt C Reconciliation</div><span>Month-End Equity Controls</span></nav>
      <main>
        <section className="hero">
          <div><h1>Excel VBA workflow, rebuilt as a web reconciliation app</h1><p>Upload raw 119 and Mgmt C workbooks, process adjusted data in memory, review exceptions, and export the final workbook.</p></div>
        </section>

        <section className="panel">
          <h2>Upload Files</h2>
          <div className="upload-grid">
            <UploadCard title="Upload RAW 119 Report" file={raw119} onChange={setRaw119} />
            <UploadCard title="Upload RAW Mgmt C Report" file={mgmtc} onChange={setMgmtc} />
          </div>
          <button className="primary" disabled={!raw119 || !mgmtc || running} onClick={runRecon}><PlayCircle />{running ? 'Processing reconciliation...' : 'Run Reconciliation'}</button>
          {error && <div className="error"><AlertCircle />{error}</div>}
          {result && <div className="success"><CheckCircle2 />Reconciliation complete. Export filename: {result.filename}</div>}
        </section>

        {result && <>
          <section className="panel dashboard">
            <div className="dashboard-head"><h2>Results Dashboard</h2><button className="download" onClick={download}><Download />Download Final Reconciliation Workbook</button></div>
            <MetricCards metrics={result.metrics} />
          </section>
          <section className="panel results">
            <div className="tabbar">{TAB_ORDER.map(tab => <button key={tab} className={activeTab === tab ? 'active' : ''} onClick={() => setActiveTab(tab)}>{tab}</button>)}</div>
            {activeTab !== 'Summary' && activeTab !== 'Approval' && <div className="tools"><label><Search size={16}/><input placeholder="Search current table" value={search} onChange={e => setSearch(e.target.value)} /></label><label className="toggle"><input type="checkbox" checked={showInternal} onChange={e => setShowInternal(e.target.checked)} />Show hidden/internal columns</label></div>}
            {activeTab === 'Summary' && <SummaryTab result={result} />}
            {activeTab === 'Approval' && <ApprovalTab />}
            {!['Summary','Approval'].includes(activeTab) && <TableView rows={result.tabs?.[activeTab] || []} search={search} showInternal={showInternal} />}
          </section>
        </>}
      </main>
    </div>
  );
}
