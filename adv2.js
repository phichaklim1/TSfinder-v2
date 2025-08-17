(function(){
  const norm = (s) => (s==null?"":String(s)).trim();
  const toNum = (v) => typeof v==="number" ? v : (typeof v==="string" ? ((n=Number(v.replace(/,/g,""))), isNaN(n)?null:n) : null);
  const escapeHtml = (s) => String(s).replace(/[&<>\"']/g,(m)=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m]));

  // แปลชื่อฟิลด์ อังกฤษ/ไทย (ต้องสะกดให้ "ตรงกับหัวคอลัมน์" ใน Excel)
const THAI_LABEL = {
  "Symbol": "เครื่องหมาย",
  "Stock symbol": "ชื่อย่อหุ้น",
  "Company Name": "ชื่อบริษัท",
  "Business details": "รายละเอียดธุรกิจ",
  "Sector": "หมวดธุรกิจ",
  "Sector2": "หมวดธุรกิจ 2",
  "Website": "เว็บไซต์",
  "%Dividend Yield": "% เงินปันผล",
  "Group": "กลุ่มหุ้น",
  "Price Zone": "โซนราคา",

  "XD in…day [Bht]": "XD ใน…วัน [บาท]",   // มีอีลิปซิส …
  "XD in...day [Bht]": "XD ใน…วัน [บาท]", // เผื่อสะกด ... ธรรมดา

  "Asset QoQ% Diff": "ทรัพย์สิน QoQ (%)",
  "Liabilities QoQ% Diff": "หนี้สิน QoQ (%)",
  "Equity QoQ% Diff": "ส่วนผู้ถือหุ้น QoQ (%)",

  "Profit or Loss Last Q": "กำไร/ขาดทุน ไตรมาสล่าสุด",
  "Profit or Loss in (M.Bht) Last Q": "กำไร/ขาดทุน (ลบ.) ไตรมาสล่าสุด",
  "EBITDA (M.Bht)": "EBITDA (ลบ.)",
  "Profit or Loss last year": "สรุปกำไร/ขาดทุนปีก่อน",

  "Cash Flow Quality": "คุณภาพกระแสเงินสด",
  "Operating Cash Flow (M.Baht)": "กระแสเงินสดจากดำเนินงาน (ลบ.)",
  "Investing Cash Flow (M.Baht)": "กระแสเงินสดจากลงทุน (ลบ.)",
  "Financing Cash Flow (M.Baht)": "กระแสเงินสดจากจัดหาเงิน (ลบ.)",
  "Net Cash Flow (M.Baht)": "กระแสเงินสดสุทธิ (ลบ.)",

  "Prev. Last (Bht)": "ราคาปิดก่อนหน้า (บาท)",
  "Open (Bht)": "ราคาเปิด (บาท)",
  "Last (Bht)": "ราคาล่าสุด (บาท)",

  "% PriceChange": "% เปลี่ยนแปลงราคา",
  "% PriceGrowth in 3 days": "% ราคาที่เติบโตใน 3 วัน",
  "TurnoverRatio(%)": "อัตราหมุนเวียน (%)",
  "% Vol Turnover": "% วอลุ่มหมุนเวียน",
  "Volume": "ปริมาณ",
  "% Vol. growth": "% การเติบโตวอลุ่ม",

  "Value (‘000 Bht)": "มูลค่า (‘000 บาท)", // เครื่องหมาย quote โค้ง
  "Value ('000 Bht)": "มูลค่า ('000 บาท)", // เผื่อใช้ quote ตรง

  "Price Spread change warning": "เตือนการเปลี่ยนแปลงช่วงราคา",
  "P/BV": "P/BV",
  "M.Cap (M.Bht)": "มูลค่าหลักทรัพย์ตามราคาตลาด (ลบ.)",
  "Equity (M.Bht)": "ส่วนของผู้ถือหุ้น (ลบ.)",
  "%FF": "% Free Float",
  "P/E": "P/E",
  "NVDR Vol.": "วอลุ่ม NVDR",
  "NVDR Volume (%Buy) since 1/6/65": "NVDR (%ซื้อ) ตั้งแต่ 1/6/65",
  "Start date": "วันที่เริ่มเข้าตลาดฯ",

  "Stock age (Y-M-D)": "อายุหุ้น (ป-ด-ว)",   // ยัติภังค์ธรรมดา
  "Stock age (Y–M–D)": "อายุหุ้น (ป–ด–ว)", // เผื่อ en dash

  "IPO price (Bht)": "ราคา IPO (บาท)",
  "Price Spread": "ช่วงราคา",
  "Securities Pledged in Margin Accounts (Vol.)": "หลักทรัพย์ค้ำบัญชีมาร์จิ้น (ปริมาณ)",
  "Stock Size S/M/L": "ขนาดหุ้น S/M/L"
};
  
  const isPercentFieldName = (h) => h.includes("%") || /(growth|change|ratio)/i.test(h);
  const isVolumeFieldName = (h) => h.toLowerCase().replace(/\./g,"").includes("volume");

  function symbolKey(headers){
    for (const k of ["Stock symbol","Symbol","stock symbol","symbol","code","ticker","ชื่อย่อหุ้น"]) if (headers.includes(k)) return k;
    return headers[0]||"";
  }
  function isNumericField(header,data){
    let num=0, tot=0;
    for(const r of data){
      const v=r[header]; if(v==null || String(v).trim()==="") continue;
      tot++; if (toNum(v)!=null) num++; if (tot>=50) break;
    }
    return tot>0 && (num/tot)>=0.7;
  }
  function formatValue(header, raw){
    const n = toNum(raw);
    if (n==null) return (raw==null || String(raw).trim()==="") ? "-" : String(raw);
    const d = isVolumeFieldName(header) ? 0 : 2;
    return new Intl.NumberFormat("en-US",{minimumFractionDigits:d, maximumFractionDigits:d}).format(n);
  }

  const rowsEl = document.getElementById("rows");
  const addRowBtn = document.getElementById("addRow");
  const clearBtn = document.getElementById("clear");
  const logicSelect = document.getElementById("logicSelect");
  const statusEl = document.getElementById("status");
  const resultTitle = document.getElementById("resultTitle");
  const tableWrap = document.getElementById("tableWrap");
  const runBottom = document.getElementById("runBottom");

  const presetName = document.getElementById("presetName");
  const savePreset = document.getElementById("savePreset");
  const presetList = document.getElementById("presetList");
  const loadPreset = document.getElementById("loadPreset");
  const deletePreset = document.getElementById("deletePreset");
  const copyLink = document.getElementById("copyLink");

  const state = {HEADERS:[], DATA:[], SYM:""};

  async function loadDB(){
    const res = await fetch("database.xlsx",{cache:"no-store"});
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(buf), { type:"array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const sheet = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:null });
    state.HEADERS = (sheet[0]||[]).map(x=>norm(x));
    state.DATA = [];
    for(let i=1;i<sheet.length;i++){
      const row = sheet[i]||[];
      if(!row.some(v=>v!=null && String(v).trim()!=="")) continue;
      const obj={}; for(let j=0;j<state.HEADERS.length;j++) obj[state.HEADERS[j]] = row[j];
      state.DATA.push(obj);
    }
    state.SYM = symbolKey(state.HEADERS);
    statusEl.textContent = "โหลดฐานข้อมูลแล้ว: " + state.DATA.length + " แถว";
    buildPresetList();
    if (!fromHash()) addRow();
  }

  function rowTemplate(idx){
    const rowId = "r"+idx, el = document.createElement("div"); el.className="adv2-row"; el.dataset.idx=idx;

    const fieldSel = document.createElement("select"); fieldSel.className="fld";
    state.HEADERS.forEach(h=>{
      const opt=document.createElement("option"); const th=THAI_LABEL[h]?` / ${THAI_LABEL[h]}`:"";
      opt.textContent = THAI_LABEL[h] ? `${THAI_LABEL[h]} / ${h}` : h;
    });

    const spacer = document.createElement("select"); spacer.className="mid"; spacer.innerHTML = `<option>—</option>`;

    const numWrap = document.createElement("div"); numWrap.className="adv2-flex numwrap adv2-hidden";
    const opSel = document.createElement("select"); opSel.innerHTML = `<option value="<">&lt;</option><option value="≤">≤</option><option value="=">=</option><option value="≥">≥</option><option value=">">&gt;</option><option value="between">ระหว่าง</option>`;
    const n1 = document.createElement("input"); n1.type="number"; n1.step="0.01"; n1.placeholder="เช่น 0.40";
    const n2 = document.createElement("input"); n2.type="number"; n2.step="0.01"; n2.placeholder="ถึง"; n2.className="n2 adv2-hidden";

    const slider = document.createElement("div"); slider.className="adv2-flex adv2-hidden sliderpack";
    slider.innerHTML = `<div class="adv2-flex" style="flex:1;gap:6px"><input type="range" class="sMin"><input type="range" class="sMax"></div>
      <div class="adv2-flex"><input type="number" class="sMinVal" step="0.01" style="width:110px"><span>ถึง</span><input type="number" class="sMaxVal" step="0.01" style="width:110px"><span class="adv2-small">%</span></div>`;

    const textWrap = document.createElement("div"); textWrap.className="adv2-flex textwrap adv2-hidden";
    const tMode = document.createElement("select"); tMode.innerHTML = `<option value="eq">เท่ากับ</option><option value="contains">มีคำว่า</option><option value="starts">ขึ้นต้นด้วย</option><option value="ends">ลงท้ายด้วย</option>`;
    const tInput = document.createElement("input"); tInput.type="text"; tInput.placeholder="ค่า/คำที่ต้องการ"; tInput.setAttribute("list","list-"+rowId);
    const dlist = document.createElement("datalist"); dlist.id="list-"+rowId;

    const kill = document.createElement("button");
    kill.textContent = "ลบบรรทัดนี้";
    kill.className = "kill adv2-kill";

    el.appendChild(fieldSel); el.appendChild(spacer); el.appendChild(numWrap); el.appendChild(textWrap); el.appendChild(kill);
    numWrap.appendChild(opSel); numWrap.appendChild(n1); numWrap.appendChild(n2); numWrap.appendChild(slider);
    textWrap.appendChild(tMode); textWrap.appendChild(tInput); el.appendChild(dlist);

   function refresh(){
  const h = fieldSel.value;
  const numeric = isNumericField(h, state.DATA);

  if (numeric){
    // โหมดตัวเลข: ซ่อนส่วนข้อความ, แสดงส่วนตัวเลข
    textWrap.classList.add("adv2-hidden");
    numWrap.classList.remove("adv2-hidden");

    // ตั้ง step: Volume = จำนวนเต็ม, อื่นๆ = ทศนิยม 2
    const isVol = h.toLowerCase().replace(/\./g,"").includes("volume");
    n1.step = n2.step = isVol ? "1" : "0.01";

    // ถ้าเป็นฟิลด์เปอร์เซ็นต์ และเลือก "ระหว่าง" → ใช้สไลเดอร์
    const pct = h && isPercentFieldName(h);
    const isBetween = (opSel.value === "between");
    const useSlider = pct && isBetween;

    // โชว์/ซ่อนสไลเดอร์
    slider.classList.toggle("adv2-hidden", !useSlider);

    if (useSlider){
      // ใช้สไลเดอร์: ซ่อนกล่องเลขคู่ (น1/น2)
      n1.classList.add("adv2-hidden");
      n2.classList.add("adv2-hidden");

      // เซ็ตช่วงของสไลเดอร์จากข้อมูลจริง
      const values = [];
      for (const r of state.DATA){
        const v = toNum(r[h]);
        if (v != null) values.push(v);
      }
      let lo = Math.min(...values), hi = Math.max(...values);
      if (!isFinite(lo)) { lo = -100; hi = 100; }

      const sMin    = slider.querySelector(".sMin");
      const sMax    = slider.querySelector(".sMax");
      const sMinVal = slider.querySelector(".sMinVal");
      const sMaxVal = slider.querySelector(".sMaxVal");

      sMin.min = sMax.min = Math.floor(lo);
      sMin.max = sMax.max = Math.ceil(hi);
      sMin.step = sMax.step = "1";

      if (sMin.value === "") sMin.value = String(Math.floor(lo));
      if (sMax.value === "") sMax.value = String(Math.ceil(hi));

      // sync ค่าระหว่างสไลเดอร์ ↔ กล่องตัวเลข ↔ น1/น2 (ที่ซ่อน)
      const syncFromSlider = ()=>{
        sMinVal.value = sMin.value;
        sMaxVal.value = sMax.value;
        n1.value = sMin.value;
        n2.value = sMax.value;
      };
      const syncFromInput = ()=>{
        sMin.value = sMinVal.value;
        sMax.value = sMaxVal.value;
        n1.value   = sMinVal.value;
        n2.value   = sMaxVal.value;
      };
      sMin.oninput = sMax.oninput = syncFromSlider;
      sMinVal.oninput = sMaxVal.oninput = syncFromInput;

      // ให้ค่าตอนเริ่มต้นตรงกัน
      syncFromSlider();

    } else {
      // ไม่ใช้สไลเดอร์: แสดงน1 และแสดง/ซ่อนน2 ตาม 'ระหว่าง'
      n1.classList.remove("adv2-hidden");
      n2.classList.toggle("adv2-hidden", !isBetween);
      slider.classList.add("adv2-hidden");
    }

  } else {
    // โหมดข้อความ: ซ่อนตัวเลขทั้งหมด แล้วโชว์ส่วนข้อความ
    numWrap.classList.add("adv2-hidden");
    slider.classList.add("adv2-hidden");
    textWrap.classList.remove("adv2-hidden");

    // เติมตัวเลือกเดาซ้ำ (datalist) ให้พิมพ์ง่าย
    const set = new Set();
    for (const r of state.DATA){
      const v = norm(r[h]);
      if (v) set.add(v);
      if (set.size > 200) break;
    }
    dlist.innerHTML = Array.from(set)
      .sort((a,b)=>a.localeCompare(b,"en",{sensitivity:"base"}))
      .map(v=>`<option value="${escapeHtml(v)}"></option>`)
      .join("");
  }
}

    opSel.addEventListener("change", refresh);
    fieldSel.addEventListener("change", refresh);
    kill.addEventListener("click", ()=> el.remove());
    refresh();
    return el;
  }

  function addRow(){ rowsEl.appendChild(rowTemplate(document.querySelectorAll(".adv2-row").length+1)); }

  function collect(){
  const arr = [];
  document.querySelectorAll(".adv2-row").forEach(row=>{
    const field = row.querySelector(".fld").value;
    const numeric = isNumericField(field, state.DATA);

    if (numeric){
      const op = row.querySelector(".numwrap select").value;

      // ค่าเริ่มต้น: เอาจากช่องตัวเลขตัวแรก (ซ้าย)
      let a = toNum(row.querySelector(".numwrap input[type=number]").value);
      let b = null;

      if (op === "between"){
        // ถ้าช่องตัวเลขตัวที่สอง (ขวา) ไม่ได้ซ่อน ใช้ค่านี้
        const n2El = row.querySelector(".numwrap .n2");
        if (n2El && !n2El.classList.contains("adv2-hidden")){
          b = toNum(n2El.value);
        } else {
          // ถ้าถูกซ่อนอยู่ แปลว่าใช้ "สไลเดอร์" → อ่านค่าจากสไลเดอร์แทน
          const sMinVal = row.querySelector(".sliderpack .sMinVal");
          const sMaxVal = row.querySelector(".sliderpack .sMaxVal");
          if (sMinVal) a = toNum(sMinVal.value);
          if (sMaxVal) b = toNum(sMaxVal.value);
        }
      }

      arr.push({kind:"num", field, op, a, b});
    } else {
      // ฟิลด์ข้อความ
      const mode = row.querySelector(".textwrap select").value;
      const text = row.querySelector(".textwrap input[type=text]").value;
      arr.push({kind:"text", field, mode, text});
    }
  });
  return arr;
}

  function matchOne(f, rec){
    const val = rec[f.field];
    if (f.kind==="num"){
      const x = toNum(val); if (x==null || f.a==null) return false;
      switch(f.op){
        case "<": return x < f.a;
        case "≤": return x <= f.a;
        case "=": return x === f.a;
        case "≥": return x >= f.a;
        case ">": return x > f.a;
        case "between":
          if (f.b==null) return false;
          const lo=Math.min(f.a,f.b), hi=Math.max(f.a,f.b);
          return x>=lo && x<=hi;
        default: return false;
      }
    }else{
      const s = norm(val), q = norm(f.text);
      if (!q) return false;
      switch(f.mode){
        case "eq": return s.localeCompare(q,"en",{sensitivity:"base"})===0;
        case "contains": return s.toLowerCase().includes(q.toLowerCase());
        case "starts": return s.toLowerCase().startsWith(q.toLowerCase());
        case "ends": return s.toLowerCase().endsWith(q.toLowerCase());
        default: return false;
      }
    }
  }

  function run(){
    const filters = collect();
    const logic = document.getElementById("logicSelect").value;
    const cols = new Set([state.SYM]); filters.forEach(f=>cols.add(f.field));
    const rows = state.DATA.filter(rec => {
      if (filters.length===0) return true;
      return logic==="AND" ? filters.every(f=>matchOne(f,rec)) : filters.some(f=>matchOne(f,rec));
    });
    render(Array.from(cols), rows);
  }

  function render(headers, recs){
    resultTitle.style.display="block"; resultTitle.textContent = `ผลลัพธ์: ${recs.length.toLocaleString("en-US")} รายการ`;
    const tbl = document.createElement("table"); tbl.className="adv2-table";
    const thead=document.createElement("thead"); const trh=document.createElement("tr");
    trh.innerHTML = headers.map(h=>`<th>${escapeHtml(THAI_LABEL[h]? `${h} / ${THAI_LABEL[h]}`: h)}</th>`).join("");
    thead.appendChild(trh); tbl.appendChild(thead);
    const tbody=document.createElement("tbody");
    recs.sort((a,b)=>String(a[state.SYM]||"").localeCompare(String(b[state.SYM]||""),"en",{sensitivity:"base"}));
    for(const r of recs){
      const tr=document.createElement("tr");
      tr.innerHTML = headers.map(h=>{
        const v=r[h]; if (toNum(v)!=null) return `<td>${escapeHtml(formatValue(h,v))}</td>`;
        return `<td>${escapeHtml(v==null?"-":String(v))}</td>`;
      }).join("");
      tbody.appendChild(tr);
    }
    tbl.appendChild(tbody); tableWrap.innerHTML=""; tableWrap.appendChild(tbl);
  }

  const LSKEY="tsf_presets_v2";
  function listPresets(){ try{return JSON.parse(localStorage.getItem(LSKEY)||"[]");}catch{ return []; } }
  function savePresets(arr){ localStorage.setItem(LSKEY, JSON.stringify(arr)); }
  function buildPresetList(){ const arr=listPresets(); presetList.innerHTML = arr.map((p,i)=>`<option value="${i}">${(p.name||('Preset '+i))}</option>`).join(""); }
  function applyConfig(cfg){
    rowsEl.innerHTML="";
    (cfg.rows||[]).forEach(()=>addRow());
    const els=[...document.querySelectorAll(".adv2-row")];
    (cfg.rows||[]).forEach((r,i)=>{
      const row=els[i];
      row.querySelector(".fld").value = r.field; row.querySelector(".fld").dispatchEvent(new Event("change"));
      if (r.kind==="num"){
        const opSel = row.querySelector(".numwrap select"); opSel.value = r.op;
        const aEl = row.querySelector(".numwrap input[type=number]"); aEl.value = r.a ?? "";
        const bEl = row.querySelector(".numwrap .n2");
        if (r.op==="between"){ bEl.classList.remove("adv2-hidden"); bEl.value = r.b ?? ""; }
        else { bEl.classList.add("adv2-hidden"); bEl.value=""; }
      }else{
        row.querySelector(".textwrap select").value = r.mode;
        row.querySelector(".textwrap input[type=text]").value = r.text || "";
      }
    });
    logicSelect.value = cfg.logic || "AND";
  }
  function cfg(){ return {logic:logicSelect.value, rows:collect()}; }
  function toHash(){ const str=btoa(unescape(encodeURIComponent(JSON.stringify(cfg())))); return "#q="+str; }
  function fromHash(){
    if (!location.hash.startsWith("#q=")) return false;
    try{
      const json = decodeURIComponent(escape(atob(location.hash.slice(3))));
      const c = JSON.parse(json); applyConfig(c); run(); return true;
    }catch(e){ console.warn("Invalid hash",e); return false; }
  }

  addRowBtn.addEventListener("click", addRow);
  runBottom.addEventListener("click", run);
  clearBtn.addEventListener("click", ()=>{ rowsEl.innerHTML=""; });
  savePreset.addEventListener("click", ()=>{
    const name = norm(presetName.value) || ("Preset "+new Date().toLocaleString());
    const arr = listPresets(); arr.push({name, ...cfg()}); savePresets(arr); buildPresetList();
  });
  loadPreset.addEventListener("click", ()=>{
    const idx = Number(presetList.value); const arr=listPresets();
    if (!isNaN(idx) && arr[idx]) applyConfig(arr[idx]);
  });
  deletePreset.addEventListener("click", ()=>{
    const idx=Number(presetList.value); const arr=listPresets();
    if (!isNaN(idx) && arr[idx]){ arr.splice(idx,1); savePresets(arr); buildPresetList(); }
  });
  copyLink.addEventListener("click", async ()=>{
    const url = location.origin + location.pathname + toHash();
    try{ await navigator.clipboard.writeText(url); statusEl.textContent="คัดลอกลิงก์แล้ว"; }catch{ statusEl.textContent="คัดลอกไม่สำเร็จ"; }
    history.replaceState(null,"", toHash());
  });

  loadDB();
})();
