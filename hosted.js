
// ---------- Globals ----------
let DATA = [];
let HEADERS = [];
let LAST_ROWS = [];
let LAST_FIELDS = [];

// ---------- Utils ----------
const $ = (sel, root=document) => root.querySelector(sel);
const $$ = (sel, root=document) => Array.from(root.querySelectorAll(sel));

function setStatus(msg){ const s = $("#status"); if(s) s.innerHTML = msg; }

function isNumberLike(field, v){
  if(v === null || v === undefined || v === "") return false;
  const x = String(v).replace(/,/g,"").trim();
  return x !== "" && !isNaN(+x);
}
function parseNumber(v){
  if(typeof v === "number") return v;
  if(v === null || v === undefined || v === "") return NaN;
  const x = String(v).replace(/,/g,"").trim();
  return x === "" ? NaN : +x;
}
function fmtCell(field, val){
  if(val === null || val === undefined || val === "") return "-";
  const name = String(field).toLowerCase();
  const zeroDecimal = ["volume", "nvdr vol", "securities pledged in margin accounts", "value (‘000 bht)", "value ('000 bht)"];
  const isVol = zeroDecimal.some(k=> name.includes(k));
  if(isNumberLike(field, val)){
    const n = parseNumber(val);
    if(isNaN(n)) return String(val);
    return isVol ? n.toLocaleString("en-US", {maximumFractionDigits:0}) : n.toLocaleString("en-US", {minimumFractionDigits:2, maximumFractionDigits:2});
  }
  // URL auto-link (for Website)
  if(/^https?:\/\//i.test(String(val))) return `<a href="${val}" target="_blank" rel="noopener">${val}</a>`;
  return String(val);
}

// ---------- Load XLSX ----------
async function loadDB(){
  try{
    if(typeof XLSX === "undefined"){ setStatus("กำลังเตรียมโมดูลอ่านไฟล์…"); setTimeout(loadDB, 300); return; }
    const resp = await fetch("database.xlsx?v="+Date.now(), {cache:"no-store"});
    if(!resp.ok){ setStatus("โหลดฐานข้อมูลไม่สำเร็จ: HTTP "+resp.status); return; }
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(buf, {type:"array"});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:""});
    if(!rows.length){ setStatus("ไฟล์ฐานข้อมูลว่าง"); return; }
    HEADERS = rows[0].map(String);
    DATA = rows.slice(1).map(r=>{ const o={}; HEADERS.forEach((h,i)=>o[h]=r[i]??""); return o;});
    setStatus("โหลดฐานข้อมูลแล้ว: "+DATA.length+" แถว");
    buildBasic();
    buildAdvanced();
  }catch(e){
    console.error(e); setStatus("โหลดฐานข้อมูลไม่สำเร็จ: "+(e.message||e));
  }
}

// ---------- Basic Search ----------
function buildBasic(){
  $("#searchBtn").addEventListener("click", doSearch);
  $("#query").addEventListener("keydown", e=>{ if(e.key==="Enter") doSearch(); });
}
function doSearch(){
  const q = $("#query").value.trim();
  if(!q){ return; }
  const key = q.toLowerCase();
  // ค้นหาในคอลัมน์ที่ชื่อเหมือน Symbol/Stock symbol
  const symHeader = HEADERS.find(h=> /^(symbol|stock symbol)$/i.test(h.trim()));
  const rows = DATA.filter(r=> String(r[symHeader]||"").toLowerCase() === key );
  const res = $("#result"); const grid = $("#grid"); const big = $("#bigtitle");
  grid.innerHTML = ""; res.style.display = "block";
  if(!rows.length){ big.textContent = q.toUpperCase()+" — ไม่พบข้อมูล"; return; }
  big.textContent = q.toUpperCase();
  // สร้างการ์ดตาม HEADERS
  HEADERS.forEach(h=>{
    const v = rows[0][h];
    const card = document.createElement("div");
    card.className = "cell";
    card.innerHTML = `<div class="cell-h">${h}</div><div class="cell-v">${fmtCell(h,v)}</div>`;
    grid.appendChild(card);
  });
}

// ---------- Advanced Filter ----------
function buildAdvanced(){
  // tab
  $$(".tab").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      $$(".tab").forEach(b=>b.classList.remove("active"));
      btn.classList.add("active");
      $$(".tabpane").forEach(p=> p.hidden = (p.id !== btn.dataset.tab));
    });
  });

  $("#addCond").addEventListener("click", addCond);
  $("#clearCond").addEventListener("click", ()=> { $("#condList").innerHTML=""; });
  $("#runFilter").addEventListener("click", runFilter);
  $("#btnCSV").addEventListener("click", exportCSV);
  $("#btnXLS").addEventListener("click", exportXLS);
  // สร้างบรรทัดแรกให้เลย
  addCond(true);
}

function addCond(first=false){
  const row = document.createElement("div");
  row.className = "cond-row";
  const fieldSel = `<select class="cond-field">${HEADERS.map(h=>`<option>${h}</option>`).join("")}</select>`;
  const opSel = `<select class="cond-op">
      <option value="=">=</option>
      <option value=">">&gt;</option>
      <option value=">=">&ge;</option>
      <option value="<">&lt;</option>
      <option value="<=">&le;</option>
      <option value="contains">มีคำว่า</option>
      <option value="between">ช่วง (เช่น 0.4-1)</option>
    </select>`;
  const valInp = `<input class="cond-val" placeholder="ค่า" />`;
  const delBtn = `<button type="button" class="btn btn-light del" ${first?'style="visibility:hidden"':""}>ลบบรรทัดนี้</button>`;
  row.innerHTML = `${fieldSel}${opSel}${valInp}${delBtn}`;
  $("#condList").appendChild(row);
  const del = $(".del", row);
  if(del) del.addEventListener("click", ()=> row.remove());
}

function runFilter(){
  const rows = DATA.slice();
  const conds = $$(".cond-row").map(row=>{
    return {
      field: $(".cond-field", row).value,
      op: $(".cond-op", row).value,
      val: $(".cond-val", row).value.trim()
    };
  });

  // ฟิลด์ที่ต้องแสดง: ชื่อหุ้น + ฟิลด์เงื่อนไข
  const symHeader = HEADERS.find(h=> /^(symbol|stock symbol)$/i.test(h.trim()));
  const fields = [symHeader, ...conds.map(c=>c.field)].filter((v,i,a)=> v && a.indexOf(v)===i);

  const pass = (r)=> conds.every(c=>{
    const raw = r[c.field];
    if(c.op === "contains"){
      return String(raw||"").toLowerCase().includes(c.val.toLowerCase());
    }
    if(c.op === "between"){
      const m = c.val.replace(/,/g,"").match(/(-?\d+(?:\.\d+)?)[\s\-~to]+(-?\d+(?:\.\d+)?)/i);
      if(!m) return false;
      const a = parseFloat(m[1]), b = parseFloat(m[2]);
      const x = parseNumber(raw);
      return !isNaN(x) && x >= Math.min(a,b) && x <= Math.max(a,b);
    }
    const x = parseNumber(raw);
    const y = parseNumber(c.val);
    if(isNaN(x) || isNaN(y)) return String(raw) == c.val;
    if(c.op==="=") return x===y;
    if(c.op===">") return x>y;
    if(c.op===">=") return x>=y;
    if(c.op==="<") return x<y;
    if(c.op==="<=") return x<=y;
    return false;
  });

  const out = rows.filter(pass).sort((a,b)=> String(a[symHeader]||"").localeCompare(String(b[symHeader]||"")));
  LAST_ROWS = out; LAST_FIELDS = fields;
  renderTable(out, fields);
}

function renderTable(rows, fields){
  const tbl = $("#filterTable");
  const ex = $("#exportBar");
  if(!rows.length){ tbl.innerHTML = "<thead><tr><th>ไม่พบข้อมูล</th></tr></thead>"; ex.hidden = true; return; }
  ex.hidden = false;
  const thead = `<thead><tr>${fields.map(h=> `<th>${h}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows.map(r=> `<tr>${fields.map(h=> `<td>${fmtCell(h, r[h])}</td>`).join("")}</tr>`).join("")}</tbody>`;
  tbl.innerHTML = thead + tbody;
}

// ---------- Export ----------
function exportCSV(){
  if(!LAST_ROWS.length) return;
  const esc = s => {
    const v = String(s).replace(/"/g,'""');
    return /[",\n]/.test(v) ? `"${v}"` : v;
  };
  const lines = [];
  lines.push(LAST_FIELDS.map(esc).join(","));
  LAST_ROWS.forEach(r=>{
    lines.push(LAST_FIELDS.map(h=> esc($(document.createElement('div')).innerHTML = fmtCell(h, r[h]), fmtCell(h, r[h])) ));
  });
  // Rebuild values without HTML tags
  const strip = html => String(html).replace(/<[^>]*>/g,"");
  const lines2 = [];
  lines2.push(LAST_FIELDS.map(esc).join(","));
  LAST_ROWS.forEach(r=>{
    lines2.push(LAST_FIELDS.map(h=> esc(strip(fmtCell(h, r[h])))).join(","));
  });
  const csv = "\ufeff" + lines2.join("\n");
  const blob = new Blob([csv], {type:"text/csv;charset=utf-8"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "filter-results.csv";
  document.body.appendChild(a); a.click(); a.remove();
}

function exportXLS(){
  if(!LAST_ROWS.length || typeof XLSX === "undefined") return exportCSV();
  const aoa = [LAST_FIELDS.slice()];
  const strip = html => String(html).replace(/<[^>]*>/g,"");
  LAST_ROWS.forEach(r=>{
    aoa.push(LAST_FIELDS.map(h=> strip(fmtCell(h, r[h])) ));
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Results");
  XLSX.writeFile(wb, "filter-results.xlsx");
}

// ---------- init ----------
document.addEventListener("DOMContentLoaded", ()=>{
  // tabs
  $$(".tab").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      $$(".tab").forEach(b=>b.classList.remove("active"));
      btn.classList.add("active");
      $$(".tabpane").forEach(p=> p.hidden = (p.id !== btn.dataset.tab));
    });
  });
  loadDB();
});
