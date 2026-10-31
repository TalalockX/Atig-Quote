/* Loads each library from the first CDN that works. */
function loadChain(urls, ready){
  return new Promise(resolve => {
    let i = 0;
    const next = () => {
      if (ready()) return resolve(true);
      if (i >= urls.length) return resolve(false);
      const s = document.createElement("script"); s.src = urls[i++];
      s.onload = () => ready() ? resolve(true) : next();
      s.onerror = next;
      document.head.appendChild(s);
    };
    next();
  });
}
const libs = Promise.all([
  loadChain(["https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js",
             "https://unpkg.com/jsqr@1.4.0/dist/jsQR.js",
             "https://cdnjs.cloudflare.com/ajax/libs/jsQR/1.4.0/jsQR.min.js"], () => !!window.jsQR),
  loadChain(["https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.js",
             "https://unpkg.com/qrcode-generator@1.4.4/qrcode.js",
             "https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js"], () => !!window.qrcode)
]);

/* Set these after deploying Code.gs. Leave SCRIPT_URL empty to run as a standalone demo. */
const CONFIG = { SCRIPT_URL: "", SECRET: "", SCANNED_BY: "Showroom tablet", SHEET_URL: "" };
/* The Connect panel at the bottom of the page saves these on the tablet, so no file editing is needed. */
try{ Object.assign(CONFIG, JSON.parse(localStorage.getItem("gqs-config") || "{}")); }catch(e){}

/* The gallery's real items (items.js, made from the Excel): serial -> [Arabic name, selling price, stock]. */
const ITEMS = window.GQS_ITEMS || {};
const itemBySerial = s => { const r = ITEMS[s]; return r ? { serial:s, en:"", ar:r[0] || s, price:r[1] || 0, cur:"SAR" } : null; };
const stockOf = serial => ITEMS[serial] ? ITEMS[serial][2] : null;
/* Units of this item on every quotation (saved or draft, never deleted ones) - the open one counted live. */
function usedQty(serial){
  let n = 0; const cur = state.quote && state.quote.id;
  Object.keys(state.saved).forEach(id => { if (id !== cur && (state.saved[id].quote.savedAt || 0) >= itemsStamp) state.saved[id].lines.forEach(l => { if (l.serial === serial) n += l.qty; }); });
  state.lines.forEach(l => { if (l.serial === serial) n += l.qty; });
  return n;
}
const leftQty = serial => stockOf(serial) === null ? null : stockOf(serial) - usedQty(serial);
function stockHtml(serial){
  const left = leftQty(serial); if (left === null) return "";
  return `<span class="stock${left < 0 ? " bad" : left === 0 ? " zero" : ""}" data-stock="${esc(serial)}">المتبقي / Left: ${left} <small>من / of ${stockOf(serial)}</small></span>`;
}
function updateStock(){
  document.querySelectorAll("#lines [data-stock]").forEach(el => { el.outerHTML = stockHtml(el.dataset.stock); });
}

const $ = id => document.getElementById(id);
const fmt = n => n.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
const r2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
const esc = s => String(s).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

const state = { quote:null, lines:[], saved:{} };
let armed = false;   // autosave stays off until the page has finished starting up

/* ---------- label format ---------- */
function parseLabel(raw){
  const p = String(raw).trim().split("|");
  if (p.length !== 5) return {error:"Not a gallery label. Expected 5 fields separated by |."};
  const [serial,en,ar,priceStr,cur] = p.map(s => s.trim());
  if (!/^[A-Z]{2,4}-\d{3,}$/.test(serial)) return {error:"Serial number looks wrong: " + serial};
  const price = Number(priceStr);
  if (!isFinite(price) || price < 0 || priceStr === "") return {error:"Price looks wrong: " + priceStr};
  if (!/^[A-Z]{3}$/.test(cur)) return {error:"Currency looks wrong: " + cur};
  if (!en) return {error:"Item name is missing."};
  return {item:{serial,en,ar,price,cur}};
}

/* ---------- quotation ---------- */
/* Quote numbers never repeat: the next one is above every saved quotation and every deleted one. */
const quoteNo = id => parseInt((String(id).match(/(\d+)$/) || [0, 0])[1], 10) || 0;   // QT-0012 and older QT-2026-0012 both give 12
function nextQuoteNo(){
  let max = 0;
  try{ max = Number(localStorage.getItem("gqs-counter")) || 0; }catch(e){}
  Object.keys(state.saved).forEach(id => { const n = quoteNo(id); if (n > max) max = n; });
  return max + 1;
}
function newQuote(){
  const n = nextQuoteNo();
  $("delConfirm").hidden = true;
  const d = new Date();
  let branch = ""; try{ branch = localStorage.getItem("gqs-branch") || ""; }catch(e){}
  state.quote = { id:"QT-" + String(n).padStart(4,"0"), customer:"", branch, custNo:"", custVat:"", custCity:"", notes:"", custPhone:"", docNo:"", validity:"",
    date:d.toISOString().slice(0,10), status:"Draft" };
  state.lines = [];
  $("cust").value = "";
  syncExtra();
  renderAll();
}
function lineTotal(l){ return r2(l.price * l.qty * (1 - l.disc/100)); }
const vatRate = () => $("vatOn").checked ? 15 : 0;
/* Per-line figures, as the Sheet shows them. Taxable amount is after discount and before VAT. */
function lineCalc(l, vatPct){
  const gross = r2(l.price * l.qty);
  const taxable = r2(gross * (1 - l.disc/100));
  const discAmt = r2(gross - taxable);
  const vatAmt = r2(taxable * vatPct / 100);
  return { gross, taxable, discAmt, vatPct, vatAmt, total:r2(taxable + vatAmt) };
}
function calc(){
  let gross=0, net=0, vat=0;
  state.lines.forEach(l => { const c = lineCalc(l, vatRate()); gross += c.gross; net += c.taxable; vat += c.vatAmt; });
  gross=r2(gross); net=r2(net); vat=r2(vat);
  return {gross, discount:r2(gross-net), net, vat, total:r2(net+vat)};
}
/* The 11 columns of the customer quotation tab, in this order. */
const LINE_COLS = ["الباركود على العلبة\nBarcode on box","اسم الصنف عربي\nItem name Arabic","رقم الفرع\nBranch no","العدد\nQty","السعر\nPrice","المبلغ\nAmount","نسبة الخصم\nDiscount %","قيمة الخصم\nDiscount value","نسبة الضريبة\nVAT %","قيمة الضريبة\nVAT value","الإجمالي\nTotal"];
function sheetRows(snap){
  return snap.lines.map(l => { const c = lineCalc(l, snap.vatPct);
    return { barcode:l.serial, name:l.ar || l.en, qty:l.qty, unitPrice:l.price, amount:c.gross,
      discPct:l.disc, discAmt:c.discAmt, vatPct:c.vatPct, vatAmt:c.vatAmt, subtotal:c.total }; });
}
function addItem(item){
  const ex = state.lines.find(l => l.serial === item.serial);
  const nm = item.ar || item.en;
  if (ex){ ex.qty += 1; renderLines(item.serial); }
  else { state.lines.push({...item, qty:1, disc:0, manualPrice: !item.price}); renderLines(item.serial); }
  renderTotals();
  const left = leftQty(item.serial);
  let msg = nm + (ex ? " added (qty " + ex.qty + ")" : " added");
  if (left !== null) msg += " - " + left + " left";
  if (left !== null && left < 0) msg += " (over stock!)";
  if (!item.price) msg += " - no selling price, enter it";
  toast(msg, 5000);
  queueAutosave();
}
/* fromCamera = true: only private labels (AQ1...) are accepted. Typed or pasted entries may also be a plain serial. */
function handleScan(raw, fromCamera){
  const t = String(raw).trim();
  if (!t) return;
  const serial = window.GQS_QR ? GQS_QR.decode(t) : null;
  if (serial){
    const it = itemBySerial(serial);
    return it ? addItem(it) : toast("Item " + serial + " is not in the item list.");
  }
  if (fromCamera) return toast("Not a gallery label.");
  if (t.includes("|")){                       // older 5-field labels still work when pasted
    const r = parseLabel(t);
    if (r.error) return toast("Not a valid label: " + r.error);
    return addItem(r.item);
  }
  const it = itemBySerial(t);
  if (!it) return toast("Item " + t.slice(0,40) + " is not in the item list.");
  addItem(it);
}

/* ---------- rendering ---------- */
/* Shows at most n rows of a table and scrolls the rest; the header stays put. */
function capRows(table, n){
  const wrap = table.closest(".tablewrap"); if (!wrap) return;
  wrap.classList.add("cap");
  const rows = table.tBodies[0] ? [...table.tBodies[0].rows].filter(r => !r.querySelector(".empty")) : [];
  if (rows.length <= n || !table.tHead || !table.tHead.offsetHeight){ wrap.style.maxHeight = ""; wrap.style.overflowY = ""; return; }
  let h = table.tHead.offsetHeight; for (let i = 0; i < n; i++) h += rows[i].offsetHeight;
  wrap.style.maxHeight = (h + 2) + "px"; wrap.style.overflowY = "auto";
}
function renderLines(flashSerial){
  $("empty").hidden = state.lines.length > 0;
  $("lines").innerHTML = state.lines.map((l,i) => `
    <tr data-i="${i}" class="${l.serial===flashSerial?"flash":""}">
      <td class="item"><b dir="auto">${esc(l.ar || l.en)}</b><span class="sub"><code>${esc(l.serial)}</code>${l.en ? `<span>${esc(l.en)}</span>` : ""}</span>${stockHtml(l.serial)}</td>
      <td class="num">${l.manualPrice ? `<div class="disc"><input type="number" inputmode="decimal" min="0" step="0.01" value="${l.price}" data-f="price" aria-label="Price"></div>` : fmt(l.price)}</td>
      <td><div class="qty">
        <button type="button" data-act="dec" aria-label="Decrease quantity">−</button>
        <input type="number" inputmode="numeric" min="1" step="1" value="${l.qty}" data-f="qty" aria-label="Quantity">
        <button type="button" data-act="inc" aria-label="Increase quantity">+</button></div></td>
      <td class="num"><div class="disc"><input type="number" inputmode="decimal" min="0" max="100" step="0.5" value="${l.disc}" data-f="disc" aria-label="Discount percent"></div></td>
      <td class="num" data-total>${fmt(lineTotal(l))}</td>
      <td><button type="button" data-act="del" aria-label="Remove item">✕</button></td>
    </tr>`).join("");
  capRows($("lines").closest("table"), 6);
  const fl = $("lines").querySelector("tr.flash"); if (fl) fl.scrollIntoView({ block:"nearest" });
}
function renderTotals(){
  const t = calc();
  $("totals").innerHTML =
    `<div><span>Subtotal</span><span>${fmt(t.gross)}</span></div>` +
    `<div><span>Discount</span><span>−${fmt(t.discount)}</span></div>` +
    `<div><span>VAT 15%</span><span>${$("vatOn").checked ? fmt(t.vat) : "off"}</span></div>` +
    `<div class="grand"><span>Total SAR</span><span>${fmt(t.total)}</span></div>`;
}
function renderHeader(){
  const q = state.quote;
  $("qid").value = q.id; $("qdate").value = q.date; $("branch").value = q.branch || "";
  $("status").textContent = q.status; $("status").className = "badge" + (q.status==="Final"?" final":"");
  const fin = q.status === "Final";
  $("finBtn").disabled = !state.lines.length;
  $("finBtn").textContent = fin ? "Set back to draft" : "Finalize";
  $("finBtn").className = fin ? "" : "primary";
  $("saveBtn").textContent = fin ? "Save changes" : "Save draft";
  $("delBtn").disabled = !(state.lines.length || state.saved[q.id]);
  renderQuoteLinks();
}
function renderSaved(){
  const ids = Object.keys(state.saved);
  $("savedPanel").hidden = !ids.length;
  $("saved").innerHTML = ids.map(id => `<button type="button" data-open="${id}">${id} · ${esc(state.saved[id].quote.customer||"No customer")} · ${state.saved[id].quote.status}</button>`).join("");
  renderPreview();
}
/* The preview shows every saved quotation, with the one being edited shown live. */
function renderPreview(){
  const map = {...state.saved};
  if (state.quote && (state.lines.length || state.saved[state.quote.id])) map[state.quote.id] = liveSnap();
  const q = Object.keys(map).map(id => map[id]);
  const head = c => "<thead><tr>" + c.map(x => `<th>${x.replace(/\n/g,"<br>")}</th>`).join("") + "</tr></thead>";
  const none = n => `<tr><td colspan="${n}" class="empty">Nothing saved yet.</td></tr>`;
  $("tQuotes").innerHTML = head(["رقم العرض\nQuote ID","العميل\nCustomer","فرع\nBranch","التاريخ\nDate","الحالة\nStatus","المجموع قبل الخصم\nSubtotal","الخصم\nDiscount","الضريبة\nVAT","الإجمالي\nTotal","أنشئ بواسطة\nCreated by"]) + "<tbody>" +
    (q.length ? q.map(s => { const t=s.totals; return `<tr><td>${s.quote.id}</td><td>${esc(s.quote.customer)}</td><td>${esc(s.quote.branch || "")}</td><td>${s.quote.date}</td><td>${s.quote.status}</td><td class="num">${fmt(t.gross)}</td><td class="num">${fmt(t.discount)}</td><td class="num">${fmt(t.vat)}</td><td class="num">${fmt(t.total)}</td><td>${esc(CONFIG.SCANNED_BY)}</td></tr>`; }).join("") : none(10)) + "</tbody>";
  capRows($("tQuotes"), 4);
  /* The open quotation's own tab: details, the 11-column table and the totals. */
  const cur = state.quote && (state.lines.length || state.saved[state.quote.id]) ? liveSnap() : null;
  $("tabTitle").textContent = cur ? "Tab " + cur.quote.id + " (this quotation)" : "This quotation's tab";
  if (!cur){ $("tabInfo").innerHTML = ""; $("tLines").innerHTML = head(LINE_COLS) + "<tbody>" + none(11) + "</tbody>"; $("tTotals").innerHTML = ""; capRows($("tLines"), 6); return; }
  const qq = cur.quote, t = cur.totals, net = t.net !== undefined ? t.net : r2(t.gross - t.discount);
  $("tabInfo").innerHTML = [["رقم العرض / Quote no.", qq.id],["العميل / Customer", qq.customer],["الفرع / Branch", qq.branch],["التاريخ / Date", qq.date],["الحالة / Status", qq.status]]
    .map(([k,v]) => `<div><span>${k}</span><b dir="auto">${esc(v || "—")}</b></div>`).join("");
  const rows = sheetRows(cur).map(r => `<tr><td>${esc(r.barcode)}</td><td dir="auto">${esc(r.name)}</td><td>${esc(qq.branch || "")}</td><td class="num">${r.qty}</td><td class="num">${fmt(r.unitPrice)}</td><td class="num">${fmt(r.amount)}</td><td class="num">${r.discPct}%</td><td class="num">${fmt(r.discAmt)}</td><td class="num">${r.vatPct}%</td><td class="num">${fmt(r.vatAmt)}</td><td class="num">${fmt(r.subtotal)}</td></tr>`);
  const tot = (label, v, cls) => `<div class="${cls || ""}"><span>${label}</span><b>${fmt(v)}</b></div>`;
  $("tLines").innerHTML = head(LINE_COLS) + "<tbody>" + (rows.length ? rows.join("") : none(11)) + "</tbody>";
  $("tTotals").innerHTML =
    tot("المجموع قبل الخصم / Subtotal", t.gross) + tot("إجمالي الخصم / Total discount", t.discount) +
    tot("المبلغ الخاضع للضريبة / Taxable amount", net) + tot("إجمالي الضريبة / Total VAT", t.vat) +
    tot("الإجمالي شامل الضريبة / Grand total", t.total, "grand");
  capRows($("tLines"), 6);
}
function renderAll(){ renderHeader(); renderLines(); renderTotals(); renderSaved(); }

/* ---------- editing ---------- */
$("lines").addEventListener("input", e => {
  const tr = e.target.closest("tr"); if (!tr || !e.target.dataset.f) return;
  const l = state.lines[+tr.dataset.i], f = e.target.dataset.f;
  let v = parseFloat(e.target.value);
  if (f === "qty") l.qty = isFinite(v) && v >= 1 ? Math.floor(v) : 1;
  else if (f === "price") l.price = isFinite(v) && v >= 0 ? r2(v) : 0;
  else l.disc = isFinite(v) ? Math.min(100, Math.max(0, v)) : 0;
  tr.querySelector("[data-total]").textContent = fmt(lineTotal(l));
  updateStock();
  renderTotals();
  queueAutosave();
});
$("lines").addEventListener("click", e => {
  const b = e.target.closest("button[data-act]"); if (!b) return;
  const i = +b.closest("tr").dataset.i, l = state.lines[i];
  if (b.dataset.act === "inc") l.qty += 1;
  if (b.dataset.act === "dec") l.qty = Math.max(1, l.qty - 1);
  if (b.dataset.act === "del"){ state.lines.splice(i,1); }
  renderLines(); renderTotals(); renderHeader();
  queueAutosave();
});
$("cust").addEventListener("input", e => { state.quote.customer = e.target.value; queueAutosave(); });
$("branch").addEventListener("change", e => {
  state.quote.branch = e.target.value;
  try{ localStorage.setItem("gqs-branch", e.target.value); }catch(err){}   // the tablet remembers its branch
  queueAutosave();
});
$("vatOn").addEventListener("change", () => { renderTotals(); queueAutosave(); });

/* ---------- save ---------- */
let pushChain = Promise.resolve(), autoT;
function liveSnap(){ return { quote:{...state.quote}, lines:state.lines.map(l => ({...l})), totals:calc(), vatPct:vatRate() }; }
function renderQuoteLinks(){
  const s = state.quote && state.saved[state.quote.id];
  const ok = !!(s && validSheetUrl(s.tabUrl));
  $("quoteLinks").hidden = !ok;
  if (ok){ $("lnkTab").href = s.tabUrl; }
}
function setSync(msg){ $("syncState").textContent = msg; }
/* Every change updates the preview straight away and is sent to the Sheet about a second later. */
function queueAutosave(){
  renderPreview(); renderHeader();
  if (!armed || (!state.lines.length && !state.saved[state.quote.id])) return;
  setSync("Unsaved changes…");
  clearTimeout(autoT);
  autoT = setTimeout(() => save(state.quote.status, { auto:true }), 900);
}
async function save(status, opts = {}){
  clearTimeout(autoT);
  const q = state.quote, auto = !!opts.auto;
  if (!state.lines.length && !auto) return toast("Add at least one item first.");
  if (!auto && status === "Final" && !q.customer.trim()) return toast("Enter the customer name before finalizing.");
  q.status = status;
  q.savedAt = Date.now();
  const snap = liveSnap();
  const prev = state.saved[q.id];
  if (prev){ snap.tabUrl = prev.tabUrl; snap.pdfUrl = prev.pdfUrl; }   // keep this quotation's Sheet links
  state.saved[q.id] = snap;
  if (auto){ renderHeader(); renderSaved(); } else renderAll();   // autosave must not redraw the lines while someone is typing
  persist();
  if (!CONFIG.SCRIPT_URL){
    setSync("Kept on this tablet only. Connect a Google Sheet below to send it.");
    if (!auto) toast(status + " saved to the preview. Connect a Google Sheet below to send it.");
    return;
  }
  setSync("Saving to Google Sheets…");
  const payload = JSON.stringify({ key:CONFIG.SECRET, createdBy:CONFIG.SCANNED_BY, quote:snap.quote, totals:snap.totals, lines:sheetRows(snap) });
  pushChain = pushChain.then(async () => {   // one at a time, in order
    try{
      const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"}, body:payload });
      const out = await res.json();
      noteSheetUrl(out);
      if (out.ok && !out.tabUrl){
        setSync("Your Apps Script is out of date: it saved nothing in the new layout. Paste the new Code.gs and deploy a New version.");
        toast("Your Apps Script is out of date. Update Code.gs.");
      } else if (out.ok){
        if (validSheetUrl(out.tabUrl) && state.saved[snap.quote.id]){
          state.saved[snap.quote.id].tabUrl = out.tabUrl; persist(); renderQuoteLinks();
        }
        const flag = out.flagged ? " " + out.flagged + " cell(s) are highlighted for review." : "";
        setSync("Saved to Google Sheets at " + new Date().toLocaleTimeString([], {hour:"2-digit", minute:"2-digit"}) + "." + flag);
        if (!auto) toast(status + " saved to Google Sheets." + flag);
      } else { setSync("The Sheet rejected the save: " + out.error); toast("Sheet rejected the save: " + out.error); }
    }catch(err){ setSync("Could not reach the Sheet. Changes are kept on this tablet."); if (!auto) toast("Could not reach the Sheet. The quotation is kept on this tablet."); }
  });
}
$("saveBtn").onclick = () => save(state.quote.status);
$("finBtn").onclick = () => save(state.quote.status === "Final" ? "Draft" : "Final");
$("newBtn").onclick = newQuote;

/* ---------- delete a quotation (tablet and Google Sheet) ---------- */
$("delBtn").onclick = () => {
  const q = state.quote;
  $("delMsg").textContent = "Delete " + q.id + (q.customer.trim() ? " (" + q.customer.trim() + ")" : "") +
    (CONFIG.SCRIPT_URL ? " and its rows in the Google Sheet?" : "?") + " This cannot be undone.";
  $("delConfirm").hidden = false;
};
$("delNo").onclick = () => { $("delConfirm").hidden = true; };
$("delYes").onclick = () => deleteQuote();
function deleteQuote(){
  const id = state.quote.id, wasSaved = !!state.saved[id];
  clearTimeout(autoT);
  delete state.saved[id];
  try{ const n = quoteNo(id), c = Number(localStorage.getItem("gqs-counter")) || 0; if (n > c) localStorage.setItem("gqs-counter", String(n)); }catch(e){}
  persist();
  newQuote();
  toast(id + " deleted");
  if (!CONFIG.SCRIPT_URL || !wasSaved){ setSync(id + " deleted from this tablet."); return; }
  setSync("Deleting " + id + " from Google Sheets…");
  const payload = JSON.stringify({ key:CONFIG.SECRET, delete:true, quoteId:id });
  pushChain = pushChain.then(async () => {
    try{
      const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"}, body:payload });
      const out = await res.json();
      noteSheetUrl(out);
      if (!out.ok) setSync(out.error === "Bad quote id"
        ? "Your Apps Script is out of date, so " + id + " was NOT removed from the Sheet. Paste the new Code.gs and deploy a New version."
        : "The Sheet could not delete " + id + ": " + out.error);
      else if (!out.tab && !out.rows && !out.summary) setSync(id + " was not found in the Sheet (nothing to remove).");
      else setSync(id + " deleted from Google Sheets" + (out.tab ? " (its tab and its row in Quotations)." : "."));
    }catch(err){ setSync("Could not reach the Sheet. " + id + " is deleted here, but its rows are still in the Sheet."); }
  });
}
$("saved").addEventListener("click", e => {
  const b = e.target.closest("[data-open]"); if (!b) return;
  const s = state.saved[b.dataset.open];
  $("delConfirm").hidden = true;
  state.quote = {...s.quote}; state.lines = s.lines.map(l => ({...l}));
  $("vatOn").checked = s.vatPct !== 0;
  $("cust").value = state.quote.customer; syncExtra(); renderAll();
  toast("Opened " + state.quote.id);
});
function persist(){ try{ localStorage.setItem("gqs-saved", JSON.stringify(state.saved)); }catch(e){} }
function restore(){ try{ const s = JSON.parse(localStorage.getItem("gqs-saved")||"{}"); if (s && typeof s==="object") state.saved = s; }catch(e){} }


/* ---------- customer quotation PDF (print layout) ---------- */
const COMPANY = {
  nameAr:"مؤسسة بدر عمر الخنبشي التجارية", addrAr:"جدة - شارع ابوبكر الصديق - الشرفية", crAr:"سجل تجاري 4030197794",
  nameEn:"Badr O. Alkhanbashi Est. for Trading", addrEn:"Jeddah, Abu Bakr as Siddiq Street", crEn:"Branch 4030197794",
  vat:"300234539600003", logo:"logo.png"
};
const EXTRA = { custNo:"custNo", custVat:"custVat", custCity:"custCity", notes:"notes", custPhone:"custPhone", docNo:"docNo", validity:"validity" };
function syncExtra(){ Object.keys(EXTRA).forEach(k => { $(EXTRA[k]).value = state.quote[k] || ""; }); }
Object.keys(EXTRA).forEach(k => $(EXTRA[k]).addEventListener("input", e => { state.quote[k] = e.target.value; queueAutosave(); }));

/* Arabic amount in words, riyals and halalas. */
function tafqeet(amount){
  const ones = ["","واحد","اثنان","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة","عشرة","أحد عشر","اثنا عشر","ثلاثة عشر","أربعة عشر","خمسة عشر","ستة عشر","سبعة عشر","ثمانية عشر","تسعة عشر"];
  const tens = ["","","عشرون","ثلاثون","أربعون","خمسون","ستون","سبعون","ثمانون","تسعون"];
  const hund = ["","مائة","مئتان","ثلاثمائة","أربعمائة","خمسمائة","ستمائة","سبعمائة","ثمانمائة","تسعمائة"];
  const u1000 = n => { const p = [], h = Math.floor(n/100), r = n % 100;
    if (h) p.push(hund[h]);
    if (r) { if (r < 20) p.push(ones[r]); else { if (r % 10) p.push(ones[r % 10]); p.push(tens[Math.floor(r/10)]); } }
    return p.join(" و"); };
  const scale = (n, one, two, few, many) => n === 1 ? one : n === 2 ? two : (n >= 3 && n <= 10) ? u1000(n) + " " + few : u1000(n) + " " + many;
  const words = n => { if (n === 0) return "صفر";
    const p = [], m = Math.floor(n / 1e6), t = Math.floor((n % 1e6) / 1000), r = n % 1000;
    if (m) p.push(scale(m, "مليون", "مليونان", "ملايين", "مليون"));
    if (t) p.push(scale(t, "ألف", "ألفان", "آلاف", "ألف"));
    if (r) p.push(u1000(r));
    return p.join(" و"); };
  const total = Math.round(amount * 100), riyal = Math.floor(total / 100), halala = total % 100;
  const parts = [];
  if (riyal || !halala) parts.push(words(riyal) + " ريال");
  if (halala) parts.push(words(halala) + " هللة");
  return parts.join(" و") + " فقط لا غير";
}

const SAR = '<svg class="sar" viewBox="0 0 1124.14 1256.39" aria-label="SAR"><path fill="currentColor" d="M699.62,1113.02h0c-20.06,44.48-33.32,92.75-38.4,143.37l424.51-90.24c20.06-44.47,33.31-92.75,38.4-143.37l-424.51,90.24Z"/><path fill="currentColor" d="M1085.73,895.8c20.06-44.47,33.32-92.75,38.4-143.37l-330.68,70.33v-135.2l292.27-62.11c20.06-44.47,33.32-92.75,38.4-143.37l-330.68,70.27V66.13c-50.67,28.45-95.67,66.32-132.25,110.99v403.35l-132.25,28.11V0c-50.67,28.44-95.67,66.32-132.25,110.99v525.69l-295.91,62.88c-20.06,44.47-33.33,92.75-38.42,143.37l334.33-71.05v170.26l-358.3,76.14c-20.06,44.47-33.32,92.75-38.4,143.37l375.04-79.7c30.53-6.35,56.77-24.4,73.83-49.24l68.78-101.97v-.02c7.14-10.55,11.3-23.27,11.3-36.97v-149.98l132.25-28.11v270.4l424.53-90.28Z"/></svg>';
const money = n => SAR + " " + fmt(n);
const mf = n => `<div class="mf"><span class="sym">${SAR}</span><span class="n">${fmt(n)}</span></div>`;
function printHtml(){
  const nowTime = new Date().toLocaleTimeString("en-US",{hour:"2-digit",minute:"2-digit",hour12:true});
  const q = state.quote, t = calc(), vp = vatRate();
  const rows = state.lines.map(l => { const c = lineCalc(l, vp);
    return `<tr class="bx">
      <td class="sub">${mf(c.total)}</td><td class="nw">${fmt(c.vatAmt)}</td><td class="nw">${fmt(c.discAmt)}</td><td class="nw">${l.disc}%</td>
      <td class="nw">${fmt(c.taxable)}</td><td class="nw">${fmt(l.price)}</td><td>${l.qty}</td><td class="pic"><img class="pic" data-serial="${esc(l.serial)}" alt=""></td>
      <td class="desc">${esc(l.ar || l.en)}</td><td class="nw">${esc(l.serial)}</td></tr>`; }).join("");
  const W = ["14%","9%","8%","5%","10%","9%","5%","17%","13%","10%"];   // widths, left to right
  const tot = [
    [t.gross, "Total (Excluding VAT)", "الإجمالي ( غير شامل ضريبة القيمة المضافة )"],
    [t.discount, "Discount", "مجموع الخصومات"],
    [t.net, "Total Taxable Amount (Excluding VAT)", "الإجمالي الخاضع للضريبة ( غير شامل ضريبة القيمة المضافة )"],
    [t.vat, "Total VAT", "مجموع ضريبة القيمة المضافة"],
    [t.total, "Total Amount Due", "إجمالي المبلغ المستحق"]
  ].map(r => `<tr><td class="v" style="width:22%">${mf(r[0])}</td><td class="en" style="width:34%">${r[1]}</td><td class="ar">${r[2]}</td></tr>`).join("");
  const header = `
  <table class="head"><tr>
    <td class="en" style="width:50%"><div class="big">${COMPANY.nameEn}</div><div>${COMPANY.addrEn}</div><div>${COMPANY.crEn}</div></td>
    <td class="ar" style="width:50%"><div class="big">${COMPANY.nameAr}</div><div>${COMPANY.addrAr}</div><div>${COMPANY.crAr}</div></td>
  </tr></table>
  <table style="margin:4px 0"><tr>
    <td style="width:38%;padding:0"><table class="kv">
      <tr><td class="en" style="width:26%">Time</td><td class="c">${nowTime}</td><td class="ar" style="width:24%"><bdi>الوقت</bdi> :</td></tr>
      <tr><td class="en">Issue Date</td><td class="c">${esc(q.date)}</td><td class="ar">التاريخ :</td></tr>
      <tr><td class="en">Doc No.</td><td class="c" style="color:#c00000">${esc(q.docNo || "")}</td><td class="ar"><bdi>رقم المستند</bdi> :</td></tr>
      <tr><td class="en">VAT NO.</td><td class="c" style="color:#0094d6">${COMPANY.vat}</td><td class="ar">الرقم الضريبي :</td></tr>
    </table></td>
    <td class="c" style="width:32%"><img src="${COMPANY.logo}" alt="" style="display:block;margin:0 auto;width:auto;height:auto;max-width:96%;max-height:27mm"></td>
    <td style="width:30%;padding:0;vertical-align:bottom"><div class="title-ar">عرض سعر</div><div class="title-en">Quotation</div></td>
  </tr></table>
  <table class="bx cu">
    <tr><td class="vn" style="width:15%">${esc(q.custNo || "")}</td><td class="cl" style="width:12%">رقم العميل<br>Customer No</td>
      <td class="v ar" colspan="3" style="width:52%">${esc(q.customer || "")}</td><td class="cla" dir="rtl" style="width:21%">اسم العميل / Customer Name</td></tr>
    <tr><td class="v" style="font-size:9px">${esc(q.custCity || "")}</td><td class="cl">البلد - المدينة<br>Country - City</td>
      <td class="v" style="width:17%">${esc(q.custVat || "")}</td><td class="cl" style="width:12%">الرقم الضريبي<br>VAT Number</td>
      <td class="v" style="width:23%">${esc(q.custPhone || "")}</td><td class="cla" dir="rtl">رقم الجوال / Mobile Number</td></tr>
  </table>
  <div class="hgap"></div>`;
  const body = `
  <table class="items bx"><colgroup>${W.map(w => `<col style="width:${w}">`).join("")}</colgroup>
    <thead>
      <tr><th class="pu">المجموع شامل الضريبة</th><th class="pu sm">ضريبة القيمة<br>المضافة <bdi>${vp}%</bdi></th><th class="pu sm" colspan="2">الخصم</th><th class="pu sm">المبلغ الخاضع<br>للضريبة</th><th class="pu">سعر الوحدة</th><th class="pu">الكمية</th><th class="ma">الصورة</th><th class="ma">الوصف</th><th class="ma">رقم الصنف</th></tr>
      <tr><th class="lb">Item Subtotal<br>(Including VAT)</th><th class="lb">Tax ${vp} %</th><th class="lb">Discount</th><th class="lb">%</th><th class="gb">Taxable Amount</th><th class="gb">Unit price</th><th class="gb">QTY</th><th class="gb">Picture</th><th class="gb">Description</th><th class="gb">Barcode</th></tr>
    </thead><tbody>${rows}</tbody></table>
  <div class="foot">
  <table class="tot cnt"><tr><td class="v" style="width:22%">${state.lines.length}</td><td class="en" style="width:34%">Total Items</td><td class="ar">إجمالي الأصناف</td></tr></table>
  <table class="tot" style="margin-top:3mm">${tot}</table>
  <div class="words ar">${tafqeet(t.total)}</div>
  <table class="notes">
    <tr><td rowspan="2" style="width:26%"><div style="height:30mm;overflow:hidden;text-align:center;padding-top:1mm"><img src="stamp.png" alt="" style="height:27mm;width:auto;max-width:94%;object-fit:contain"></div></td>
      <td class="ar" style="width:60%;height:18mm">${esc(q.notes || "")}</td><td class="ar c" style="width:14%"><bdi>ملاحظات</bdi><br>Notes</td></tr>
    <tr><td class="ar" style="height:10mm">مدة ( <b style="display:inline-block;min-width:12mm;text-align:center">${esc(q.validity || "")}</b> ) يوم / أيام</td><td class="ar c"><bdi>صلاحية العرض</bdi><br>Validity</td></tr>
  </table></div>`;
  return `<table class="page"><thead><tr><td>${header}</td></tr></thead><tbody><tr><td>${body}</td></tr></tbody></table>`;
}
$("pdfBtn").addEventListener("click", () => {
  if (!state.lines.length) return toast("Add at least one item first.");
  $("printSheet").innerHTML = printHtml();
  /* Arabic stays bold; English words (labels, headings, address) are set in normal weight. Values stay bold. */
  const tw = document.createTreeWalker($("printSheet"), NodeFilter.SHOW_TEXT), nodes = [];
  while (tw.nextNode()) nodes.push(tw.currentNode);
  nodes.forEach(n => {
    if (!/[A-Za-z]/.test(n.nodeValue) || n.parentElement.closest(".items tbody, .v, .vn, .c, .stamp, .notes")) return;
    const span = document.createElement("span");
    span.innerHTML = esc(n.nodeValue).replace(/[A-Za-z(][A-Za-z0-9 .,()\-%'&]*/g, m => '<span class="nb">' + m + '</span>');
    n.replaceWith(...span.childNodes);
  });
  document.body.classList.add("pq");
  const done = () => { document.body.classList.remove("pq"); window.removeEventListener("afterprint", done); };
  window.addEventListener("afterprint", done);
  /* Item pictures: items/<serial>.jpg (also .jpeg .png .webp). Missing pictures leave the space empty. */
  const exts = ["jpg","jpeg","png","webp"];
  const settle = im => new Promise(res => {
    let k = 0; const serial = encodeURIComponent(im.dataset.serial);
    im.onload = () => res(); 
    im.onerror = () => { k++; if (k < exts.length) im.src = "items/" + serial + "." + exts[k]; else { im.style.display = "none"; res(); } };
    im.src = "items/" + serial + "." + exts[0];
  });
  const fixed = [...$("printSheet").querySelectorAll("img:not(.pic)")].map(im => im.complete ? 0 : new Promise(r => { im.onload = r; im.onerror = r; }));
  const pics = [...$("printSheet").querySelectorAll("img.pic")].map(settle);
  Promise.race([Promise.all([...fixed, ...pics]), new Promise(r => setTimeout(r, 15000))]).then(() => setTimeout(() => window.print(), 60));
});

/* ---------- toast ---------- */
let tt, tf;
function toast(m, ms = 3200){
  const t = $("toast"); t.textContent = m; t.classList.remove("out"); t.hidden = false;
  clearTimeout(tt); clearTimeout(tf);
  tf = setTimeout(() => t.classList.add("out"), ms - 600);   // fade over the last 0.6 s
  tt = setTimeout(() => { t.hidden = true; t.classList.remove("out"); }, ms);
}

/* ---------- search the item list ---------- */
const norm = s => String(s).toLowerCase().replace(/[\u064B-\u0652\u0640]/g,"").replace(/[أإآ]/g,"ا").replace(/ة/g,"ه").replace(/ى/g,"ي");
let searchIndex = null;
function searchItems(q){
  if (!searchIndex) searchIndex = Object.keys(ITEMS).map(s => [s, norm(ITEMS[s][0]), s.toLowerCase()]);
  const toks = norm(q).split(/\s+/).filter(Boolean); if (!toks.length) return [];
  const out = [];
  for (const [s, n, sl] of searchIndex){
    if (toks.every(t => n.includes(t) || sl.includes(t))){ out.push(s); if (out.length >= 12) break; }
  }
  return out;
}
$("searchIn").addEventListener("input", () => {
  const list = searchItems($("searchIn").value);
  $("searchOut").innerHTML = list.map(s => { const r = ITEMS[s], left = leftQty(s);
    return `<button type="button" class="hit" data-s="${esc(s)}"><b dir="auto">${esc(r[0] || s)}</b><span><code>${esc(s)}</code> · ${r[1] ? fmt(r[1]) + " SAR" : "no price"} · المتبقي / Left: ${left}</span></button>`; }).join("")
    || ($("searchIn").value.trim() ? '<p class="hint">No item found.</p>' : "");
});
$("searchOut").addEventListener("click", e => {
  const b = e.target.closest("[data-s]"); if (!b) return;
  addItem(itemBySerial(b.dataset.s)); $("searchIn").value = ""; $("searchOut").innerHTML = "";
});
$("pasteForm").addEventListener("submit", e => {
  e.preventDefault();
  const v = $("searchIn").value.trim(); if (!v) return;
  const direct = (window.GQS_QR && GQS_QR.decode(v)) || ITEMS[v] || v.includes("|");   // a scanned/typed code or exact serial
  if (direct) handleScan(v);
  else { const list = searchItems(v); if (list.length) addItem(itemBySerial(list[0])); else toast("No item found for \"" + v.slice(0,40) + "\"."); }
  $("searchIn").value = ""; $("searchOut").innerHTML = "";
});


/* ---------- update items from an Excel file (shared through the Google Sheet) ---------- */
const ITEMS_KEY = "gqs-items";
let itemsStamp = 0, pending = null;
function setItems(map, stamp){
  Object.keys(ITEMS).forEach(k => delete ITEMS[k]); Object.assign(ITEMS, map);
  searchIndex = null; itemsStamp = stamp || 0;
}
function saveItemsLocal(){ try{ localStorage.setItem(ITEMS_KEY, JSON.stringify({ updatedAt:itemsStamp, items:ITEMS })); }catch(e){} }
try{ const s = JSON.parse(localStorage.getItem(ITEMS_KEY) || "null"); if (s && s.items && Object.keys(s.items).length) setItems(s.items, s.updatedAt || 1); }catch(e){}
function showItemsInfo(extra){
  const n = Object.keys(ITEMS).length;
  $("itemsInfo").textContent = n + " items" + (itemsStamp > 1 ? " · updated " + new Date(itemsStamp).toLocaleString() : " · built-in list") + (extra ? " · " + extra : "");
}
const toNum = v => { const n = Number(String(v == null ? "" : v).replace(/[^\d.\-]/g, "")); return isFinite(n) ? n : 0; };
async function readExcel(file){
  const ok = await loadChain(["https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
                              "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
                              "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js"], () => !!window.XLSX);
  if (!ok) throw new Error("The Excel reader did not load. Check the internet connection and try again.");
  const wb = XLSX.read(await file.arrayBuffer(), { type:"array" });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, raw:false, defval:"" });
  const hi = rows.findIndex(r => r.some(c => String(c).trim() === "الباركود"));
  if (hi < 0) throw new Error('Could not find the column "الباركود" in the first sheet.');
  const col = {}; rows[hi].forEach((c, i) => { if (String(c).trim()) col[String(c).trim()] = i; });
  for (const need of ["الباركود", "الصنف", "سعر البيع", "العدد"]) if (!(need in col)) throw new Error('Missing column: ' + need);
  const map = {}; let skipped = 0, noPrice = 0, noName = 0, inStock = 0;
  rows.slice(hi + 1).forEach(r => {
    const s = String(r[col["الباركود"]]).trim();
    if (!s) { skipped++; return; }
    const name = String(r[col["الصنف"]]).trim(), price = Math.round(toNum(r[col["سعر البيع"]]) * 100) / 100, qty = Math.round(toNum(r[col["العدد"]]));
    if (!price) noPrice++; if (!name) noName++; if (qty > 0) inStock++;
    map[s] = [name, price, qty];
  });
  const count = Object.keys(map).length;
  if (!count) throw new Error("No items found in that file.");
  return { map, count, skipped, noPrice, noName, inStock };
}
async function pushItems(map){
  const items = Object.keys(map).map(s => [s, map[s][0], map[s][1], map[s][2]]);
  const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"}, body:JSON.stringify({ key:CONFIG.SECRET, putItems:true, items }) });
  const out = await res.json();
  if (!out.ok) throw new Error(out.error === "Bad quote id" ? "Your Apps Script is out of date. Paste the new Code.gs and deploy a New version." : out.error);
  return out.updatedAt;
}
$("itemsFile").addEventListener("change", async e => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  $("itemsApply").hidden = $("itemsCancel").hidden = true; pending = null;
  $("itemsMsg").textContent = "Reading " + f.name + "…";
  try{
    pending = await readExcel(f);
    const p = pending;
    $("itemsMsg").textContent = p.count + " items found (" + p.inStock + " in stock" + (p.noPrice ? ", " + p.noPrice + " without a price" : "") + (p.noName ? ", " + p.noName + " without a name" : "") + (p.skipped ? "; " + p.skipped + " rows without a barcode skipped" : "") + "). This replaces the current list. Apply?";
    $("itemsApply").hidden = $("itemsCancel").hidden = false;
  }catch(err){ $("itemsMsg").textContent = err.message || "Could not read that file."; }
});
$("itemsCancel").onclick = () => { pending = null; $("itemsApply").hidden = $("itemsCancel").hidden = true; $("itemsMsg").textContent = "Cancelled. The item list was not changed."; };
$("itemsApply").onclick = async () => {
  if (!pending) return;
  const map = pending.map; pending = null; $("itemsApply").hidden = $("itemsCancel").hidden = true;
  setItems(map, Date.now()); saveItemsLocal(); renderLines(); showItemsInfo();
  if (!CONFIG.SCRIPT_URL){ $("itemsMsg").textContent = "Item list updated on this tablet. Connect the Google Sheet to share it with the other tablets."; return; }
  $("itemsMsg").textContent = "Updated here. Sharing with the other tablets…";
  try{
    const t = await pushItems(map); itemsStamp = t; saveItemsLocal(); showItemsInfo();
    $("itemsMsg").textContent = "Item list updated and shared. The other tablets get it next time they open the page.";
  }catch(err){ $("itemsMsg").textContent = "Updated on this tablet only. Could not share it: " + (err.message || err); }
};
/* Other tablets pick up the newest shared list when the page opens. */
async function pullItems(){
  if (!CONFIG.SCRIPT_URL) return;
  try{
    const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"}, body:JSON.stringify({ key:CONFIG.SECRET, getItems:true, since:itemsStamp }) });
    const out = await res.json();
    if (!out.ok || out.unchanged || !out.items || !out.items.length || out.updatedAt <= itemsStamp) return;
    const map = {}; out.items.forEach(r => { map[r[0]] = [r[1], r[2], r[3]]; });
    setItems(map, out.updatedAt); saveItemsLocal(); renderAll(); showItemsInfo("from the Google Sheet");
  }catch(e){}
}


/* ---------- decoding ---------- */
/* Uses the browser's built-in QR reader when it has one (Chrome, Edge, Android), otherwise jsQR. */
let detector = null;
try{ if ("BarcodeDetector" in window) detector = new BarcodeDetector({ formats:["qr_code"] }); }catch(e){}
async function decodeCanvas(canvas, tries){
  if (detector){
    try{ const r = await detector.detect(canvas); if (r.length && r[0].rawValue) return r[0].rawValue; }catch(e){}
  }
  if (window.jsQR){
    const w = canvas.width, h = canvas.height;
    const d = canvas.getContext("2d",{willReadFrequently:true}).getImageData(0,0,w,h);
    const res = jsQR(d.data, w, h, { inversionAttempts: tries });
    if (res){
      try{ return res.binaryData ? new TextDecoder("utf-8").decode(new Uint8Array(res.binaryData)) : res.data; }
      catch(e){ return res.data; }
    }
  }
  return null;
}
const engineName = () => detector ? "built-in reader + jsQR" : (window.jsQR ? "jsQR" : null);
async function selfTest(){
  const el = $("engine"), eng = engineName();
  if (!eng){ el.textContent = "Scanner engine did not load. Check the internet connection and reload the page."; return; }
  try{
    const serial = GQS_QR.encode(Object.keys(ITEMS)[0] || "12345");
    qrcode.stringToBytes = qrcode.stringToBytesFuncs["UTF-8"];
    const qr = qrcode(0,"M"); qr.addData(serial,"Byte"); qr.make();
    const img = new Image(); img.src = qr.createDataURL(4,4); await img.decode();
    const c = document.createElement("canvas"), s = 3; c.width = img.naturalWidth*s; c.height = img.naturalHeight*s;
    const x = c.getContext("2d",{willReadFrequently:true}); x.imageSmoothingEnabled = false;
    x.fillStyle = "#fff"; x.fillRect(0,0,c.width,c.height); x.drawImage(img,0,0,c.width,c.height);
    const t = await decodeCanvas(c, "attemptBoth");
    el.textContent = "Scanner: " + eng + ". Self-test " + (t === serial ? "passed." : "failed (read: " + t + ").") + " " + Object.keys(ITEMS).length + " items loaded.";
  }catch(e){ el.textContent = "Scanner: " + eng + ". " + Object.keys(ITEMS).length + " items loaded."; }
}

/* live camera */
let stream=null, scanning=false, lastCode="", lastAt=0;
const video=$("video"), work=document.createElement("canvas"), wctx=work.getContext("2d",{willReadFrequently:true});
async function startCam(){
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) return camFail();
  try{
    stream = await navigator.mediaDevices.getUserMedia({ video:{ facingMode:{ideal:"environment"} }, audio:false });
    video.srcObject = stream; await video.play();
    video.hidden=false; $("frame").hidden=false; $("vfmsg").hidden=true;
    $("camBtn").textContent="Stop camera"; scanning=true; tick();
  }catch(e){ camFail(); }
}
function camFail(){ $("vfmsg").textContent="The camera is not available in this view. Use Scan from photo instead."; $("vfmsg").hidden=false; toast("Camera blocked or unavailable. Try Scan from photo."); }
function stopCam(){
  scanning=false; if (stream) stream.getTracks().forEach(t=>t.stop()); stream=null;
  video.hidden=true; $("frame").hidden=true; $("vfmsg").hidden=false; $("vfmsg").textContent="Camera is off. Start the camera, or take a photo of a label.";
  $("camBtn").textContent="Start camera";
}
let busy = false;
async function tick(){
  if (!scanning) return;
  if (!busy && video.readyState >= 2 && video.videoWidth){
    busy = true;
    const s = Math.min(1, 960 / video.videoWidth);
    work.width = Math.round(video.videoWidth*s); work.height = Math.round(video.videoHeight*s);
    wctx.drawImage(video,0,0,work.width,work.height);
    const txt = await decodeCanvas(work, "dontInvert");
    busy = false;
    const now = Date.now();
    if (txt && (txt !== lastCode || now - lastAt > 2000)){ lastCode = txt; lastAt = now; handleScan(txt, true); if (navigator.vibrate) navigator.vibrate(60); }
  }
  requestAnimationFrame(tick);
}
$("camBtn").onclick = () => scanning ? stopCam() : startCam();

/* photo fallback */
$("photoBtn").onclick = () => $("photoIn").click();
$("photoIn").addEventListener("change", async e => {
  const f = e.target.files[0]; e.target.value=""; if (!f) return;
  if (!engineName()) return toast("Scanner engine did not load. Check the internet connection and reload the page.");
  try{
    toast("Reading photo…");
    const bmp = await createImageBitmap(f);
    let txt = null;
    // Try several sizes: very large photos and very small ones both fail at a single scale.
    for (const max of [1400, 900, 2000, 600, 400]){
      const s = max / Math.max(bmp.width, bmp.height);
      work.width = Math.max(1, Math.round(bmp.width*s)); work.height = Math.max(1, Math.round(bmp.height*s));
      wctx.fillStyle = "#fff"; wctx.fillRect(0,0,work.width,work.height);
      wctx.drawImage(bmp,0,0,work.width,work.height);
      txt = await decodeCanvas(work, "attemptBoth");
      if (txt) break;
    }
    txt ? handleScan(txt, true) : toast("No QR code found in that photo. Fill the frame with the label and avoid glare.");
  }catch(err){ toast("Could not read that image."); }
});

libs.then(() => { if (window.qrcode) selfTest(); else $("engine").textContent = Object.keys(ITEMS).length + " items loaded."; });
restore(); newQuote(); showItemsInfo(); pullItems();
armed = true; renderPreview();

/* ---------- link to the Sheet ---------- */
const validSheetUrl = u => /^https:\/\/docs\.google\.com\/spreadsheets\/d\/[\w-]+/.test(u || "");
function updateSheetLink(){
  document.querySelectorAll(".sheetLink").forEach(a => {
    const ok = validSheetUrl(CONFIG.SHEET_URL);
    a.hidden = !ok; if (ok) a.href = CONFIG.SHEET_URL;
  });
}
/* The script tells us the Sheet's address in its answers; remember it so staff can open the Sheet. */
function noteSheetUrl(out){
  if (out && validSheetUrl(out.sheetUrl) && out.sheetUrl !== CONFIG.SHEET_URL){
    CONFIG.SHEET_URL = out.sheetUrl;
    try{ localStorage.setItem("gqs-config", JSON.stringify(CONFIG)); }catch(e){}
    $("cfgSheet").value = CONFIG.SHEET_URL; updateSheetLink();
  }
}

/* ---------- connect to Google Sheets ---------- */
function showConfig(msg){
  $("cfgSheet").value = CONFIG.SHEET_URL; updateSheetLink();
  $("cfgUrl").value = CONFIG.SCRIPT_URL; $("cfgKey").value = CONFIG.SECRET; $("cfgName").value = CONFIG.SCANNED_BY;
  $("cfgState").textContent = msg || (CONFIG.SCRIPT_URL ? "Saved on this tablet. Press Test connection." : "Not connected. Quotations are kept on this tablet only.");
}
$("cfgSave").onclick = () => {
  CONFIG.SCRIPT_URL = $("cfgUrl").value.trim(); CONFIG.SECRET = $("cfgKey").value.trim();
  CONFIG.SCANNED_BY = $("cfgName").value.trim() || "Showroom tablet";
  const typed = $("cfgSheet").value.trim(); CONFIG.SHEET_URL = validSheetUrl(typed) ? typed : "";
  try{ localStorage.setItem("gqs-config", JSON.stringify(CONFIG)); }catch(e){ return toast("This browser would not save the settings."); }
  showConfig(); renderSaved(); toast("Connection settings saved on this tablet.");
};
$("cfgTest").onclick = async () => {
  if (!$("cfgUrl").value.trim()) return toast("Paste the Apps Script web app URL first.");
  $("cfgSave").onclick();
  $("cfgState").textContent = "Testing…";
  try{
    const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"},
      body:JSON.stringify({ key:CONFIG.SECRET, ping:true }) });
    const out = await res.json();
    noteSheetUrl(out);
    showConfig(!out.ok ? "The Sheet answered: " + out.error + ". Check the secret."
      : (out.features || []).includes("items") ? "Connected. Saves will go to your Google Sheet. Script version " + out.version + "."
      : "Connected, but your Apps Script is out of date (it does not know the latest features). Paste the new Code.gs and deploy a New version.");
  }catch(err){ showConfig("Could not reach the script. Check the URL, and that access is set to Anyone."); }
};
showConfig();
