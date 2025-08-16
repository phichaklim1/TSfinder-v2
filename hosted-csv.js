
let DATA=[], HEADERS=[], LAST_ROWS=[], LAST_FIELDS=[];
const $ = (s,r=document)=>r.querySelector(s); const $$ = (s,r=document)=>Array.from(r.querySelectorAll(s));
function setStatus(m){ const s=$("#status"); if(s) s.innerHTML=m; }

function parseCSV(text){
  const rows=[]; let row=[], val="", inQ=false;
  for(let i=0;i<text.length;i++){
    const c=text[i], n=text[i+1];
    if(inQ){
      if(c==='"' && n==='"'){ val+='"'; i++; }
      else if(c==='"'){ inQ=false; }
      else val+=c;
    }else{
      if(c==='"'){ inQ=true; }
      else if(c===','){ row.push(val); val=""; }
      else if(c==='\n'){ row.push(val); rows.push(row); row=[]; val=""; }
      else if(c==='\r'){}
      else val+=c;
    }
  }
  row.push(val); rows.push(row);
  if(rows.length && rows[rows.length-1].every(x=>x==="")) rows.pop();
  return rows;
}

function isNumberLike(v){ if(v===null||v===undefined||v==="") return false; const x=String(v).replace(/,/g,"").trim(); return x!=="" && !isNaN(+x); }
function parseNumber(v){ if(typeof v==="number") return v; if(v==null||v==="") return NaN; const x=String(v).replace(/,/g,"").trim(); return x===""?NaN:+x; }
function fmtCell(field, val){
  if(val==null||val==="") return "-";
  const name = String(field).toLowerCase();
  const zero = ["volume","nvdr vol","securities pledged in margin accounts","value (‘000 bht)","value ('000 bht)"];
  const isVol = zero.some(k=>name.includes(k));
  if(isNumberLike(val)){
    const n=parseNumber(val); if(isNaN(n)) return String(val);
    return isVol ? n.toLocaleString("en-US",{maximumFractionDigits:0})
                 : n.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
  }
  if(/^https?:\/\//i.test(String(val))) return `<a href="${val}" target="_blank" rel="noopener">${val}</a>`;
  return String(val);
}

async function loadDB(){
  try{
    const resp = await fetch("database.csv?v="+Date.now(), {cache:"no-store"});
    if(!resp.ok){ setStatus("โหลดฐานข้อมูลไม่สำเร็จ: HTTP "+resp.status); return; }
    const text = await resp.text();
    const rows = parseCSV(text);
    if(!rows.length){ setStatus("ไฟล์ว่าง"); return; }
    HEADERS = rows[0];
    DATA = rows.slice(1).map(r=>{ const o={}; HEADERS.forEach((h,i)=>o[h]=r[i]??""); return o; });
    setStatus("โหลดฐานข้อมูลแล้ว: "+DATA.length+" แถว");
    build();
  }catch(e){ setStatus("โหลดฐานข้อมูลไม่สำเร็จ: "+(e.message||e)); }
}

function build(){
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
  addCond(false);
}

function doSearch(){
  const q=$("#query").value.trim().toLowerCase();
  if(!q) return;
  const sym = HEADERS.find(h=>/^(symbol|stock symbol)$/i.test(h.trim())) || HEADERS[0];
  const rows = DATA.filter(r=> String(r[sym]||"").toLowerCase()===q );
  const res=$("#result"), grid=$("#grid"), big=$("#bigtitle");
  grid.innerHTML=""; res.style.display="block";
  if(!rows.length){ big.textContent=q.toUpperCase()+" — ไม่พบข้อมูล"; return; }
  big.textContent=q.toUpperCase();
  HEADERS.forEach(h=>{
    const v=rows[0][h]; const d=document.createElement("div");
    d.className="cell"; d.innerHTML=`<div class="cell-h">${h}</div><div class="cell-v">${fmtCell(h,v)}</div>`;
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
  const conds = $$(".cond-row").map(r=>({field:$(".cond-field",r).value, op:$(".cond-op",r).value, val:$(".cond-val",r).value.trim()}));
  const sym = HEADERS.find(h=>/^(symbol|stock symbol)$/i.test(h.trim())) || HEADERS[0];
  const fields=[sym, ...conds.map(c=>c.field)].filter((v,i,a)=>a.indexOf(v)===i);
  const pass = r => conds.every(c=>{
    const raw=r[c.field];
    if(c.op==="contains") return String(raw||"").toLowerCase().includes(c.val.toLowerCase());
    if(c.op==="between"){
      const m=c.val.replace(/,/g,"").match(/(-?\d+(?:\.\d+)?)[\s\-~to]+(-?\d+(?:\.\d+)?)/i);
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
  const tbody = `<tbody>${rows.map(r=>`<tr>${fields.map(h=>`<td>${fmtCell(h,r[h])}</td>`).join("")}</tr>`).join("")}</tbody>`;
  tbl.innerHTML = thead + tbody;
}

function exportCSV(){
  if(!LAST_ROWS.length) return;
  const esc=s=>{ const v=String(s).replace(/"/g,'""'); return /[",\n]/.test(v)?`"${v}"`:v; };
  const strip=html=>String(html).replace(/<[^>]*>/g,"");
  const lines=[]; lines.push(LAST_FIELDS.map(esc).join(","));
  LAST_ROWS.forEach(r=> lines.push(LAST_FIELDS.map(h=> esc(strip(fmtCell(h,r[h])))).join(",")) );
  const blob=new Blob(["\ufeff"+lines.join("\n")], {type:"text/csv;charset=utf-8"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="filter-results.csv"; a.click();
}

document.addEventListener("DOMContentLoaded", loadDB);
