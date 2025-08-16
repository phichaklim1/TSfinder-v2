
let DATA=[], HEADERS=[], LAST_ROWS=[], LAST_FIELDS=[];
const $=(s,r=document)=>r.querySelector(s); const $$=(s,r=document)=>Array.from(r.querySelectorAll(s));
function setStatus(m){ const el=$("#status"); if(el) el.textContent=m; }
function isNumberLike(v){ if(v===null||v===undefined||v==="") return false; const x=String(v).replace(/,/g,"").trim(); return x!=="" && !isNaN(+x); }
function parseNumber(v){ if(typeof v==="number") return v; if(v==null||v==="") return NaN; const x=String(v).replace(/,/g,"").trim(); return x===""?NaN:+x; }
function fmt(field,val){
  if(val==null||val==="") return "-";
  const name=String(field).toLowerCase();
  const intCols=["volume","value (‘000 bht)","value ('000 bht)","securities pledged in margin accounts"];
  const isInt=intCols.some(k=>name.includes(k));
  if(isNumberLike(val)){
    const n=parseNumber(val); if(isNaN(n)) return String(val);
    return isInt? n.toLocaleString("en-US",{maximumFractionDigits:0}) : n.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
  }
  if(/^https?:\/\//i.test(String(val))) return `<a href="${val}" target="_blank" rel="noopener">${val}</a>`;
  return String(val);
}

async function loadHosted(){
  try{
    const resp = await fetch("database.xlsx?v="+Date.now(), {cache:"no-store"});
    if(!resp.ok) throw new Error("HTTP "+resp.status);
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(buf), {type:"array"});
    const sh = wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sh], {defval:""});
    if(!rows.length) throw new Error("ไฟล์ว่าง");
    HEADERS = Object.keys(rows[0]);
    DATA = rows;
    setStatus?.("โหลดฐานข้อมูลแล้ว: "+DATA.length+" แถว");
    buildUI();
  }catch(e){
    setStatus?.("โหลดฐานข้อมูลไม่สำเร็จ: "+(e.message||e));
  }
}

function loadLocal(){
  const picker=$("#picker");
  picker.addEventListener("change", async (ev)=>{
    const f=ev.target.files?.[0]; if(!f) return;
    const buf=await f.arrayBuffer();
    const wb=XLSX.read(new Uint8Array(buf),{type:"array"});
    const sh=wb.SheetNames[0];
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[sh], {defval:""});
    HEADERS=rows.length?Object.keys(rows[0]):[]; DATA=rows;
    const st=$("#status"); if(st) st.textContent="โหลดข้อมูลสำเร็จ: "+DATA.length+" แถว";
  });
  buildUI();
}

function buildUI(){
  $$(".tab").forEach(btn=>btn.addEventListener("click",()=>{
    $$(".tab").forEach(b=>b.classList.remove("active")); btn.classList.add("active");
    $$(".tabpane").forEach(p=> p.hidden=(p.id!==btn.dataset.tab));
  }));
  $("#searchBtn").addEventListener("click", doSearch);
  $("#query").addEventListener("keydown", e=>{ if(e.key==="Enter") doSearch(); });
  $("#addCond").addEventListener("click", ()=> addCond(true));
  $("#clearCond").addEventListener("click", ()=> $("#condList").innerHTML="");
  $("#runFilter").addEventListener("click", runFilter);
  $("#btnCSV").addEventListener("click", exportCSV);
  $("#btnXLS").addEventListener("click", exportXLS);
  addCond(false);
}

function doSearch(){
  const q=$("#query").value.trim().toLowerCase();
  if(!q){ return; }
  const sym = HEADERS.find(h=>/^(symbol|stock symbol)$/i.test(h.trim())) || HEADERS[0];
  const rows = DATA.filter(r=> String(r[sym]||"").toLowerCase()===q );
  const res=$("#result"), grid=$("#grid"), big=$("#bigtitle");
  grid.innerHTML=""; res.style.display="block";
  if(!rows.length){ big.textContent=q.toUpperCase()+" — ไม่พบข้อมูล"; return; }
  big.textContent=q.toUpperCase();
  const row = rows[0];
  HEADERS.forEach(h=>{
    const d=document.createElement("div"); d.className="cell";
    d.innerHTML = `<div class="cell-h">${h}</div><div class="cell-v">${fmt(h,row[h])}</div>`;
    grid.appendChild(d);
  });
}

function addCond(first){
  const row=document.createElement("div"); row.className="cond-row";
  const fieldSel=`<select class="cond-field">${HEADERS.map(h=>`<option>${h}</option>`).join("")}</select>`;
  const opSel=`<select class="cond-op">
    <option value="=">=</option><option value=">">&gt;</option>
    <option value=">=">&ge;</option><option value="<">&lt;</option>
    <option value="<=">&le;</option><option value="contains">มีคำว่า</option>
    <option value="between">ช่วง (เช่น 0.4-1)</option></select>`;
  const valInput=`<input class="cond-val" placeholder="ค่า">`;
  const delBtn=`<button type="button" class="btn btn-light del" ${first?'style="visibility:hidden"':""}>ลบบรรทัดนี้</button>`;
  row.innerHTML = `${fieldSel}${opSel}${valInput}${delBtn}`;
  $("#condList").appendChild(row);
  const del=$(".del",row); if(del) del.addEventListener("click", ()=> row.remove());
}

function runFilter(){
  if(!DATA.length) return;
  const conds = $$(".cond-row").map(r=>({field:$(".cond-field",r).value, op:$(".cond-op",r).value, val:$(".cond-val",r).value.trim()}));
  const sym = HEADERS.find(h=>/^(symbol|stock symbol)$/i.test(h.trim())) || HEADERS[0];
  const fields=[sym, ...conds.map(c=>c.field)].filter((v,i,a)=>a.indexOf(v)===i);
  const pass = r => conds.every(c=>{
    const raw=r[c.field];
    if(c.op==="contains") return String(raw||"").toLowerCase().includes(c.val.toLowerCase());
    if(c.op==="between"){
      const m=c.val.replace(/,/g,"").match(/(-?\\d+(?:\\.\\d+)?)[\\s\\-~to]+(-?\\d+(?:\\.\\d+)?)/i);
      if(!m) return false; const a=+m[1], b=+m[2]; const x=parseNumber(raw);
      return !isNaN(x) && x>=Math.min(a,b) && x<=Math.max(a,b);
    }
    const x=parseNumber(raw), y=parseNumber(c.val);
    if(isNaN(x) || isNaN(y)) return String(raw)==c.val;
    if(c.op==="=") return x===y; if(c.op===">") return x>y; if(c.op===">=") return x>=y;
    if(c.op==="<") return x<y; if(c.op==="<=") return x<=y; return false;
  });
  const out = DATA.filter(pass).sort((a,b)=> String(a[sym]||"").localeCompare(String(b[sym]||"")));
  LAST_ROWS=out; LAST_FIELDS=fields; renderTable(out, fields);
}

function renderTable(rows, fields){
  const tbl=$("#filterTable"), ex=$("#exportBar");
  if(!rows.length){ tbl.innerHTML="<thead><tr><th>ไม่พบข้อมูล</th></tr></thead>"; ex.hidden=true; return; }
  ex.hidden=false;
  const thead = `<thead><tr>${fields.map(h=>`<th>${h}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows.map(r=>`<tr>${fields.map(h=>`<td>${fmt(h,r[h])}</td>`).join("")}</tr>`).join("")}</tbody>`;
  tbl.innerHTML = thead + tbody;
}

function exportCSV(){
  if(!LAST_ROWS.length) return;
  const esc=s=>{ const v=String(s).replace(/"/g,'""'); return /[",\\n]/.test(v)?`"${v}"`:v; };
  const strip=html=>String(html).replace(/<[^>]*>/g,"");
  const lines=[]; lines.push(LAST_FIELDS.map(esc).join(","));
  LAST_ROWS.forEach(r=> lines.push(LAST_FIELDS.map(h=> esc(strip(fmt(h,r[h])))).join(",")) );
  const blob=new Blob(["\\ufeff"+lines.join("\\n")], {type:"text/csv;charset=utf-8"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="filter-results.csv"; a.click();
}

function exportXLS(){
  if(!LAST_ROWS.length) return;
  const rows=[LAST_FIELDS, ...LAST_ROWS.map(r=> LAST_FIELDS.map(h=> String(r[h]??"")))];
  const ws=XLSX.utils.aoa_to_sheet(rows); const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "results"); XLSX.writeFile(wb, "filter-results.xlsx");
}

document.addEventListener("DOMContentLoaded", ()=>{
  if(window.LOCAL_MODE){ loadLocal(); } else { loadHosted(); }
});
