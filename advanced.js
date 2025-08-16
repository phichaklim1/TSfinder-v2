(function () {
  // ===== Utilities =====
  const norm = (s) => (s == null ? "" : String(s)).trim();
  const toNum = (v) =>
    typeof v === "number"
      ? v
      : typeof v === "string"
      ? ((n = Number(v.replace(/,/g, ""))), isNaN(n) ? null : n)
      : null;

  // หา header สำหรับสัญลักษณ์หุ้น
  const symbolHeader = (HEADERS) => {
    for (const k of ["Stock symbol", "Symbol", "stock symbol", "symbol", "code", "ticker", "ชื่อย่อหุ้น"]) {
      if (HEADERS.includes(k)) return k;
    }
    return HEADERS[0] || "";
  };

  // แปลชื่อฟิลด์ อังกฤษ/ไทย
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
    "XD in…day [Bht]": "XD ใน…วัน [บาท]",
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
    "Value (‘000 Bht)": "มูลค่า (‘000 บาท)",
    "Price Spread change warning": "เตือนการเปลี่ยนแปลงช่วงราคา",
    "P/BV": " ",
    "M.Cap (M.Bht)": "มูลค่าหลักทรัพย์ตามราคาตลาด (ลบ.)",
    "Equity (M.Bht)": "ส่วนของผู้ถือหุ้น (ลบ.)",
    "%FF": "% Free Float",
    "P/E": "P/E",
    "NVDR Vol.": "วอลุ่ม NVDR",
    "NVDR Volume (%Buy) since 1/6/65": "NVDR (%ซื้อ) ตั้งแต่ 1/6/65",
    "Start date": "วันที่เริ่มเข้าตลาดฯ",
    "Stock age (Y-M-D)": "อายุหุ้น (ป-ด-ว)",
    "IPO price (Bht)": "ราคา IPO (บาท)",
    "Price Spread": "ช่วงราคา",
    "Securities Pledged in Margin Accounts (Vol.)": "หลักทรัพย์ค้ำบัญชีมาร์จิ้น (ปริมาณ)"
  };

  // คำพ้อง Price Zone
  const ZONE_SYNONYMES = {
    lower: new Set(["LOWER", "LOW", "LOW ZONE", "LOWER ZONE", "โซนล่าง", "ล่าง"]),
    middle: new Set(["MIDDLE", "MID", "MID ZONE", "MIDDLE ZONE", "โซนกลาง", "กลาง"]),
    upper: new Set(["UPPER", "HIGH", "HIGH ZONE", "UPPER ZONE", "โซนบน", "บน"]),
  };
  const isZoneMatch = (cell, want) => {
    if (!cell) return false;
    const s = String(cell).trim().toUpperCase();
    if (want === "lower") return ZONE_SYNONYMES.lower.has(s);
    if (want === "middle") return ZONE_SYNONYMES.middle.has(s);
    if (want === "upper") return ZONE_SYNONYMES.upper.has(s);
    return false;
  };

  // ฟอร์แมตตัวเลข: 2 ตำแหน่ง ยกเว้น Volume ไม่มีทศนิยม
  const toFixedDisplay = (header, raw) => {
    const n = toNum(raw);
    if (n == null) return (raw == null || String(raw).trim() === "") ? "-" : String(raw);
    const isVolume = header.toLowerCase().replace(/\./g, "").includes("volume");
    const d = isVolume ? 0 : 2;
    return new Intl.NumberFormat("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }).format(n);
  };

  // ===== DOM =====
  const fieldSelect = document.getElementById("fieldSelect");
  const valueSelect = document.getElementById("valueSelect");
  const categoricalControls = document.getElementById("categoricalControls");
  const numericControls = document.getElementById("numericControls");
  const opSelect = document.getElementById("opSelect");
  const num1 = document.getElementById("num1");
  const num2 = document.getElementById("num2");
  const runSearch = document.getElementById("runSearch");
  const statusEl = document.getElementById("status");
  const result = document.getElementById("result");
  const resultTitle = document.getElementById("resultTitle");
  const tableWrap = document.getElementById("tableWrap");

  let HEADERS = [], DATA = [], SYMBOL_KEY = "";

  async function loadDatabase() {
    try {
      const res = await fetch("database.xlsx", { cache: "no-store" });
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const sheet = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
      HEADERS = (sheet[0] || []).map(x => norm(x));
      DATA = [];
      for (let i = 1; i < sheet.length; i++) {
        const row = sheet[i] || [];
        if (!row.some(v => v != null && String(v).trim() !== "")) continue;
        const obj = {};
        for (let j = 0; j < HEADERS.length; j++) obj[HEADERS[j]] = row[j];
        DATA.push(obj);
      }
      SYMBOL_KEY = symbolHeader(HEADERS);
      statusEl.textContent = "โหลดฐานข้อมูลแล้ว: " + DATA.length + " แถว";

      buildFieldOptions();
      onFieldChange();
    } catch (err) {
      statusEl.textContent = "โหลดฐานข้อมูลไม่สำเร็จ: " + err.message;
    }
  }

  function buildFieldOptions() {
    fieldSelect.innerHTML = "";
    const frag = document.createDocumentFragment();
    HEADERS.forEach(h => {
      if (!h) return;
      const opt = document.createElement("option");
      const th = THAI_LABEL[h] ? ` / ${THAI_LABEL[h]}` : "";
      opt.value = h;
      opt.textContent = `${h}${th}`;
      frag.appendChild(opt);
    });
    fieldSelect.appendChild(frag);
  }

  function isNumericField(header) {
    // ถ้าค่าเป็นตัวเลข >= 70% ของตัวอย่างที่ไม่ว่าง -> numeric
    let num = 0, tot = 0;
    for (const r of DATA) {
      const v = r[header];
      if (v == null || String(v).trim() === "") continue;
      tot++;
      if (toNum(v) != null) num++;
      if (tot >= 50) break;
    }
    return tot > 0 && (num / tot) >= 0.7;
  }

  function onFieldChange() {
    const header = fieldSelect.value;
    const numeric = isNumericField(header);

    const isVol = header.toLowerCase().replace(/\./g, "").includes("volume");
    num1.step = isVol ? "1" : "0.01";
    num2.step = isVol ? "1" : "0.01";

    if (numeric) {
      categoricalControls.style.display = "none";
      numericControls.style.display = "flex";
      num2.style.display = opSelect.value === "between" ? "inline-block" : "none";
      valueSelect.innerHTML = "";
    } else {
      categoricalControls.style.display = "flex";
      numericControls.style.display = "none";
      const values = new Set();
      for (const r of DATA) {
        const v = r[header];
        const s = norm(v);
        if (s) values.add(s);
      }
      const list = Array.from(values).sort((a, b) => a.localeCompare(b, "en", { sensitivity: "base" }));
      valueSelect.innerHTML = "";
      const frag = document.createDocumentFragment();

      if (header === "Price Zone") {
        const zones = [
          { key: "lower", text: "โซนล่าง (Lower)" },
          { key: "middle", text: "โซนกลาง (Middle)" },
          { key: "upper", text: "โซนบน (Upper)" },
        ];
        zones.forEach(z => {
          const opt = document.createElement("option");
          opt.value = "__zone__:" + z.key;
          opt.textContent = z.text;
          frag.appendChild(opt);
        });
        const sep = document.createElement("option");
        sep.disabled = true; sep.textContent = "────────";
        frag.appendChild(sep);
      }

      list.forEach(v => {
        const opt = document.createElement("option");
        opt.value = v;
        opt.textContent = v;
        frag.appendChild(opt);
      });
      valueSelect.appendChild(frag);
    }
  }

  opSelect.addEventListener("change", () => {
    num2.style.display = opSelect.value === "between" ? "inline-block" : "none";
  });
  fieldSelect.addEventListener("change", onFieldChange);

  function runFilter() {
    const header = fieldSelect.value;
    const numeric = isNumericField(header);

    let hits = [];
    if (!numeric) {
      const sel = valueSelect.value;
      if (header === "Price Zone" && sel.startsWith("__zone__:")) {
        const key = sel.split(":")[1];
        hits = DATA.filter(r => isZoneMatch(r[header], key));
      } else {
        const want = norm(sel).toUpperCase();
        hits = DATA.filter(r => norm(r[header]).toUpperCase() === want);
      }
    } else {
      const op = opSelect.value;
      const a = toNum(num1.value);
      const b = toNum(num2.value);
      hits = DATA.filter(r => {
        const x = toNum(r[header]);
        if (x == null || a == null) return false;
        switch (op) {
          case "<": return x < a;
          case "≤": return x <= a;
          case "=": return x === a;
          case "≥": return x >= a;
          case ">": return x > a;
          case "between":
            if (b == null) return false;
            const lo = Math.min(a, b), hi = Math.max(a, b);
            return x >= lo && x <= hi;
          default: return false;
        }
      });
    }

    renderTable(header, hits);
  }

  function renderTable(header, rows) {
    result.style.display = "block";
    const thLabel = THAI_LABEL[header] ? `${header} / ${THAI_LABEL[header]}` : header;

    resultTitle.textContent = `ผลลัพธ์: ${rows.length.toLocaleString("en-US")} รายการ`;

    const tbl = document.createElement("table");
    tbl.className = "data-table";
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    trh.innerHTML = `<th>Symbol / ชื่อย่อหุ้น</th><th>${escapeHtml(thLabel)}</th>`;
    thead.appendChild(trh);
    tbl.appendChild(thead);

    const tbody = document.createElement("tbody");
    const symKey = SYMBOL_KEY;

    rows.sort((a, b) =>
      String(a[symKey] || "").localeCompare(String(b[symKey] || ""), "en", { sensitivity: "base" })
    );

    for (const r of rows) {
      const tr = document.createElement("tr");
      const sym = r[symKey] == null ? "-" : String(r[symKey]);
      const val = toFixedDisplay(header, r[header]);
      tr.innerHTML = `<td>${escapeHtml(sym)}</td><td>${escapeHtml(val)}</td>`;
      tbody.appendChild(tr);
    }
    tbl.appendChild(tbody);

    tableWrap.innerHTML = "";
    tableWrap.appendChild(tbl);
  }

  const escapeHtml = (s) =>
    String(s).replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));

  runSearch.addEventListener("click", runFilter);
  [num1, num2].forEach(inp => inp.addEventListener("keydown", e => { if (e.key === "Enter") runFilter(); }));

  loadDatabase();
})();
