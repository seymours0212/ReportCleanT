import React, { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

const fieldKeywords = {
  date: ['日期', '銷售日期', '交易日期'],
  sales: ['業務', '業務員', '銷售員'],
  customer: ['客戶', '客戶名稱', '客戶簡稱'],
  productCode: ['產品代號', '品號', '料號'],
  productName: ['產品名稱', '品名', '品項'],
  qty: ['數量', '銷售數量'],
  unitPrice: ['單價', '售價'],
  amount: ['金額', '銷售金額', '未稅金額'],
  profit: ['毛利', '毛利額'],
  region: ['區域', '地區']
};

const diagnosisRows = [
  ['P01', '表首', '公司名稱、報表名稱、報表期間等資訊混在資料區前方。', '容易被誤判為資料列。', '標記為表首，不納入分析。', '第一筆明細應從欄位名稱列之後開始。', 'Excel / Power Query / Python'],
  ['P02', '表尾', '系統說明、列印說明或其他結尾文字。', '可能被誤判為交易資料。', '標記為表尾，保留來源但排除分析。', '最後一筆明細不應是說明文字。', 'Excel / Power Query'],
  ['P03', '頁首', '頁碼、列印人員、報表期間可能重複出現在中段。', '會中斷資料表結構。', '標記為頁首並排除於分析資料。', '最終資料不得含頁碼或列印資訊。', 'Excel / Power Query'],
  ['P04', '頁尾', '每頁底部可能有簽核欄、備註或說明。', '會造成空白或文字資料列。', '標記為頁尾，保留檢查但不納入明細。', '分析資料不含非交易文字。', 'Excel / Power Query'],
  ['P05', '空白列', '報表為了閱讀與列印加入空白資料列。', '樞紐分析表會出現空白項目。', '標記為空白列，版本B排除。', '版本B空白列數應為0。', 'Excel / Power Query'],
  ['P06', '群組標題', '例如「業務：A001 林業務」或分類段落標題。', '可能被當成交易列。', '解析為分類來源，不直接納入明細。', '群組標題不得出現在版本B。', 'Excel / Power Query / Python'],
  ['P07', '小計', '各業務或分類後方有小計列。', '若直接分析會重複加總。', '版本A保留檢核，版本B排除。', '明細合計需等於小計。', 'Excel / Power Query'],
  ['P08', '總計', '報表最後有總計列。', '若納入分析會重複計算。', '版本A保留檢核，版本B排除。', '版本B合計需可與總計比對。', 'Excel / Power Query'],
  ['P09', '分類值缺漏', '業務、客戶、區域常用留白表示同上。', '無法依分類彙總。', '向下填滿，但不可跨越不同群組。', '每筆明細均有分類值。', 'Excel / Power Query'],
  ['P10', '欄位名稱', '欄名不一定在第一列，也可能名稱不標準。', '匯入後欄位錯位。', '用關鍵字偵測真正欄位列。', '欄位數與欄位名稱應一致。', 'Excel / Power Query / Python'],
  ['P11', '日期型態', '日期可能是文字或混合格式。', '無法依年月排序與分析。', '轉換為日期型態。', '可依月份排序與分組。', 'Excel / Power Query'],
  ['P12', '金額型態', '金額可能含千分位逗號或文字符號。', '無法加總或計算毛利率。', '移除逗號並轉數值。', 'SUM結果需正確。', 'Excel / Power Query / Python'],
  ['P13', '數量型態', '數量、單價、毛利可能為文字。', '會造成平均與加總錯誤。', '統一轉數值型態。', '可正常加總與平均。', 'Excel / Power Query'],
  ['P14', '業務代號與姓名混合', '例如「A001 林業務」。', '不利建立維度資料。', '拆成業務代號與業務姓名。', '每列均有代號或姓名。', 'Excel / Power Query / Python'],
  ['P15', '明細列辨識', '明細、小計、總計、標題混在同一區域。', '自動化清理容易誤刪或誤留。', '建立資料列類型欄位。', '抽樣檢查資料列類型。', 'Excel / Power Query / Python'],
  ['P16', '資料品質風險', '留白、格式不一致、欄位變動都會影響結果。', '清理後數字可能失真。', '保留問題說明與清理狀態。', '筆數、金額、分類完整性需檢核。', 'Excel / Power Query / Python']
].map(([id, type, desc, impact, action, check, tools]) => ({ id, type, desc, impact, action, check, tools }));

function normalizeText(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/,/g, '').replace(/[$NTD元 ]/g, '').trim();
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function toDateValue(value) {
  if (!value) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'number') {
    const d = XLSX.SSF.parse_date_code(value);
    if (d) return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  const s = normalizeText(value).replace(/\//g, '-');
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? s : d.toISOString().slice(0, 10);
}

function splitSales(value) {
  const text = normalizeText(value).replace(/^業務[:：]?\s*/, '');
  const m = text.match(/^([A-Za-z]+\d+|\d+|[A-Za-z0-9_-]+)\s*(.*)$/);
  if (!m) return { salesCode: '', salesName: text };
  return { salesCode: m[1] || '', salesName: (m[2] || '').trim() };
}

function getCell(row, index) {
  return index >= 0 ? row[index] : '';
}

function findHeaderRow(rows) {
  let best = { index: -1, score: 0 };
  rows.forEach((row, idx) => {
    const text = row.map(normalizeText).join('|');
    let score = 0;
    Object.values(fieldKeywords).forEach(keys => {
      if (keys.some(k => text.includes(k))) score += 1;
    });
    if (score > best.score) best = { index: idx, score };
  });
  return best.score >= 3 ? best.index : -1;
}

function detectColumns(header) {
  const result = {};
  Object.entries(fieldKeywords).forEach(([key, words]) => {
    result[key] = header.findIndex(h => words.some(w => normalizeText(h).includes(w)));
  });
  return result;
}

function classifyRow(row, rowIndex, headerIndex, cols) {
  const text = row.map(normalizeText).filter(Boolean).join(' ');
  const nonEmpty = row.filter(v => normalizeText(v) !== '').length;
  if (nonEmpty === 0) return '空白列';
  if (rowIndex < headerIndex) return '表首';
  if (rowIndex === headerIndex) return '欄位名稱';
  if (/總計|合計|Grand Total/i.test(text)) return '總計';
  if (/小計|Subtotal/i.test(text)) return '小計';
  if (/頁次|頁碼|列印|報表期間|製表|系統|備註|注意/i.test(text)) return rowIndex > headerIndex ? '頁首' : '表首';
  if (/^業務[:：]/.test(text) || /^區域[:：]/.test(text) || /^客戶[:：]/.test(text)) return '群組標題';
  const hasDate = normalizeText(getCell(row, cols.date)) !== '';
  const hasProduct = normalizeText(getCell(row, cols.productCode)) !== '' || normalizeText(getCell(row, cols.productName)) !== '';
  const hasAmount = toNumber(getCell(row, cols.amount)) !== null;
  if ((hasDate && hasProduct) || (hasProduct && hasAmount)) return '明細';
  if (rowIndex > headerIndex && nonEmpty <= 2) return '表尾';
  return '無法判斷';
}

function parseWorkbook(file, callback) {
  const reader = new FileReader();
  reader.onload = e => {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    callback(rows);
  };
  reader.readAsArrayBuffer(file);
}

function cleanRows(rows) {
  const headerIndex = findHeaderRow(rows);
  if (headerIndex < 0) throw new Error('找不到欄位名稱列，請確認報表是否包含日期、客戶、產品、數量、金額等欄位。');
  const header = rows[headerIndex].map((h, i) => normalizeText(h) || `欄位${i + 1}`);
  const cols = detectColumns(header);
  const normalized = [];
  const rawPreview = rows.slice(0, 30).map((r, i) => ({ '#': i + 1, ...Object.fromEntries(r.map((v, j) => [`C${j + 1}`, v])) }));
  let current = { sales: '', customer: '', region: '' };

  rows.forEach((row, idx) => {
    const type = classifyRow(row, idx, headerIndex, cols);
    const salesRaw = normalizeText(getCell(row, cols.sales));
    const customerRaw = normalizeText(getCell(row, cols.customer));
    const regionRaw = normalizeText(getCell(row, cols.region));
    const joined = row.map(normalizeText).join(' ');

    if (type === '群組標題') {
      if (/業務[:：]/.test(joined)) current.sales = joined.replace(/^.*?業務[:：]\s*/, '').trim();
      if (/區域[:：]/.test(joined)) current.region = joined.replace(/^.*?區域[:：]\s*/, '').trim();
      if (/客戶[:：]/.test(joined)) current.customer = joined.replace(/^.*?客戶[:：]\s*/, '').trim();
    }
    if (type === '明細') {
      if (salesRaw) current.sales = salesRaw;
      if (customerRaw) current.customer = customerRaw;
      if (regionRaw) current.region = regionRaw;
    }

    const salesValue = salesRaw || current.sales;
    const customerValue = customerRaw || current.customer;
    const regionValue = regionRaw || current.region;
    const split = splitSales(salesValue);
    const amount = toNumber(getCell(row, cols.amount));
    const profit = toNumber(getCell(row, cols.profit));
    const qty = toNumber(getCell(row, cols.qty));
    const unitPrice = toNumber(getCell(row, cols.unitPrice));

    const issue = [];
    if (type === '明細' && !salesValue) issue.push('業務缺漏');
    if (type === '明細' && amount === null) issue.push('銷售金額非數值或缺漏');
    if (type === '明細' && !getCell(row, cols.date)) issue.push('日期缺漏');

    normalized.push({
      原始列號: idx + 1,
      資料列類型: type,
      日期: type === '明細' ? toDateValue(getCell(row, cols.date)) : '',
      業務: salesValue,
      業務代號: split.salesCode,
      業務姓名: split.salesName,
      客戶名稱: type === '明細' ? customerValue : '',
      產品代號: type === '明細' ? normalizeText(getCell(row, cols.productCode)) : '',
      產品名稱: type === '明細' ? normalizeText(getCell(row, cols.productName)) : '',
      數量: qty,
      單價: unitPrice,
      銷售金額: amount,
      毛利: profit,
      區域: type === '明細' ? regionValue : '',
      清理狀態: type === '明細' ? '可分析' : type === '小計' || type === '總計' ? '檢核保留' : '排除分析',
      問題說明: issue.join('；') || (type === '無法判斷' ? '需人工確認' : '')
    });
  });

  const versionA = normalized.filter(r => ['明細', '小計', '總計'].includes(r.資料列類型));
  const versionB = normalized.filter(r => r.資料列類型 === '明細');
  return { headerIndex, header, cols, normalized, rawPreview, versionA, versionB };
}

function formatNumber(n) {
  return Number(n || 0).toLocaleString('zh-TW');
}

function makeFileName(base) {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  const stamp = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  return `${base.replace(/\.xlsx$/i, '')}_資料清洗_${stamp}.xlsx`;
}

export default function App() {
  const [fileName, setFileName] = useState('');
  const [result, setResult] = useState(null);
  const [error, setError] = useState('');
  const [activeTab, setActiveTab] = useState('versionB');
  const [dragging, setDragging] = useState(false);

  const kpis = useMemo(() => {
    if (!result) return [];
    const detail = result.versionB;
    const sum = key => detail.reduce((a, r) => a + (Number(r[key]) || 0), 0);
    const unique = key => new Set(detail.map(r => r[key]).filter(Boolean)).size;
    const countType = type => result.normalized.filter(r => r.資料列類型 === type).length;
    const avgUnit = detail.length ? sum('單價') / detail.filter(r => Number(r.單價)).length : 0;
    return [
      ['原始資料列數', result.normalized.length],
      ['清理後明細列數', detail.length],
      ['小計列數', countType('小計')],
      ['總計列數', countType('總計')],
      ['空白列數', countType('空白列')],
      ['無法判斷列數', countType('無法判斷')],
      ['銷售金額合計', formatNumber(sum('銷售金額'))],
      ['毛利合計', formatNumber(sum('毛利'))],
      ['平均單價', formatNumber(Math.round(avgUnit || 0))],
      ['客戶數', unique('客戶名稱')],
      ['產品數', unique('產品代號') || unique('產品名稱')],
      ['業務人數', unique('業務代號') || unique('業務姓名')]
    ];
  }, [result]);

  function handleFile(file) {
    if (!file) return;
    setError('');
    setFileName(file.name);
    parseWorkbook(file, rows => {
      try {
        setResult(cleanRows(rows));
        setActiveTab('versionB');
      } catch (err) {
        setError(err.message);
        setResult(null);
      }
    });
  }

  function download(kind) {
    if (!result) return;
    const wb = XLSX.utils.book_new();
    const base = fileName || 'ERP銷售日報.xlsx';
    if (kind === 'A' || kind === 'all') XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.versionA), '版本A_含小計總計');
    if (kind === 'B' || kind === 'all') XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.versionB), '版本B_純明細');
    if (kind === 'all') {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.normalized), '完整清理紀錄');
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(diagnosisRows), '報表問題診斷表');
    }
    XLSX.writeFile(wb, makeFileName(base));
  }

  const previewData = result ? (activeTab === 'raw' ? result.rawPreview : activeTab === 'versionA' ? result.versionA : result.versionB).slice(0, 50) : [];
  const columns = previewData[0] ? Object.keys(previewData[0]) : [];

  return (
    <div className="app">
      <header className="hero">
        <div>
          <p className="eyebrow">先診斷，再清理；先保留檢核依據，再產出分析資料</p>
          <h1>ERP 銷售日報清理工具</h1>
          <p>將主管列印報表轉換為可分析的標準資料表，支援小計/總計檢核與純明細輸出。</p>
        </div>
        <div className="hero-card">
          <strong>輸出版本</strong>
          <span>版本A：含小計與總計</span>
          <span>版本B：純明細資料</span>
        </div>
      </header>

      <section className={`upload ${dragging ? 'dragging' : ''}`} onDragOver={e => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={e => { e.preventDefault(); setDragging(false); handleFile(e.dataTransfer.files[0]); }}>
        <div>
          <h2>上傳 Excel 報表(.xlsx)</h2>
          <p>{fileName || '拖曳檔案到此，或點選右側按鈕選擇檔案'}</p>
          {error && <p className="error">{error}</p>}
        </div>
        <label className="button primary">
          選擇檔案
          <input hidden type="file" accept=".xlsx,.xls" onChange={e => handleFile(e.target.files[0])} />
        </label>
      </section>

      {result && <>
        <section className="kpi-grid">
          {kpis.map(([label, value]) => <div className="kpi" key={label}><span>{label}</span><strong>{value}</strong></div>)}
        </section>

        <section className="panel actions">
          <div>
            <h2>下載清理結果</h2>
            <p>版本A用於對帳檢核，版本B用於樞紐分析表與 Power BI。</p>
          </div>
          <div className="button-group">
            <button onClick={() => download('A')}>下載版本A</button>
            <button onClick={() => download('B')}>下載版本B</button>
            <button className="primary" onClick={() => download('all')}>下載完整結果</button>
          </div>
        </section>

        <section className="panel">
          <h2>資料預覽</h2>
          <div className="tabs">
            <button className={activeTab === 'raw' ? 'active' : ''} onClick={() => setActiveTab('raw')}>原始資料預覽</button>
            <button className={activeTab === 'versionA' ? 'active' : ''} onClick={() => setActiveTab('versionA')}>版本A：含小計與總計</button>
            <button className={activeTab === 'versionB' ? 'active' : ''} onClick={() => setActiveTab('versionB')}>版本B：純明細</button>
          </div>
          <DataTable columns={columns} rows={previewData} />
        </section>
      </>}

      <section className="panel">
        <h2>報表問題診斷表</h2>
        <DataTable columns={['id', 'type', 'desc', 'impact', 'action', 'check', 'tools']} rows={diagnosisRows} labels={{ id: '問題編號', type: '問題類型', desc: '問題描述', impact: '對分析的影響', action: '建議處理方式', check: '檢核方式', tools: '後續可用工具' }} />
      </section>
    </div>
  );
}

function DataTable({ columns, rows, labels = {} }) {
  if (!rows.length) return <p className="empty">尚無資料。</p>;
  return <div className="table-wrap"><table><thead><tr>{columns.map(c => <th key={c}>{labels[c] || c}</th>)}</tr></thead><tbody>{rows.map((row, i) => <tr key={i}>{columns.map(c => <td key={c}>{row[c] ?? ''}</td>)}</tr>)}</tbody></table></div>;
}
