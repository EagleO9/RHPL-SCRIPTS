/************************************************************
 * FINAL UPDATED Code.gs (ORDER QTY & REMAINING LOGIC)
 * DESCRIPTION COLUMN FETCHED FROM ORDER SHEET (COL F)
 * UPDATED: Item Name & Description separated by SPACE (Hyphen removed)
 * FIX: Re-print PDF fetching logic from INWARD sheet via Challan Number
 * FIX: Partial Order Logic in getOrderDetails (CASE-INSENSITIVE + ITEM & DESC)
 * UPDATE: Avg/PCS Removed & Contractor Name Conditional
 * GROUPING FIX: Item + Description + Contractor in Re-print
 ************************************************************/

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("form");
}

/* ===========================
    FORMATTERS
=========================== */
function formatPDFDate_(d){
  try{
    const dt = new Date(d);
    const opts = { day: "numeric", month: "short", year: "numeric" };
    return dt.toLocaleDateString("en-GB", opts).replace(",", "");
  }catch(e){ return d; }
}

function formatPDFTime_(v){
  try{
    if (!v) return v;
    if (Object.prototype.toString.call(v) === "[object Date]") {
      return Utilities.formatDate(v, "Asia/Kolkata", "HH:mm:ss");
    }
    const m = String(v).match(/(\d{2}:\d{2}:\d{2})/);
    return m ? m[1] : v;
  }catch(e){ return v; }
}

function fmt2(v){
  const n = Number(v);
  return isNaN(n) ? "0.00" : n.toFixed(2);
}

/* ===========================
    DROPDOWN LIST LOADERS
=========================== */
function getItemList() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Master");
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  return [...new Set(sh.getRange(2,1,last-1,1).getValues().flat().filter(Boolean))];
}

function getChallanList(limit) {
  const sh = SpreadsheetApp.getActive().getSheetByName("INWARD");
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 6, last - 1, 1).getValues().flat(); 
  const out = [];
  const seen = new Set();
  for (let i = vals.length - 1; i >= 0; i--) {
    const v = String(vals[i]).trim();
    if (v && !seen.has(v)) {
      seen.add(v);
      out.push(v);
      if (limit && out.length >= limit) break;
    }
  }
  return out;
}

function getOrderList() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Order");
  if (!sh) return []; 
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 1).getValues().flat();
  return [...new Set(vals.filter(Boolean))];
}

/* ===========================
    getOrderDetails (UPDATED FOR CASE-INSENSITIVE ITEM + DESC)
    Returns pending items only based on Item Name + Description match. 
=========================== */
function getOrderDetails(orderId) {
  const ss = SpreadsheetApp.getActive();
  const orderSh = ss.getSheetByName("Order");
  const inwardSh = ss.getSheetByName("INWARD");
  
  if (!orderSh) return { found: false };
  
  const lastOrderRow = orderSh.getLastRow();
  if (lastOrderRow < 2) return { found: false };
  
  // 1. Calculate packed counts with Case-Insensitive Composite Key (Item|Desc)
  const packedCounts = {};
  if (inwardSh) {
    const lastInwardRow = inwardSh.getLastRow();
    if (lastInwardRow >= 2) {
      const inwardData = inwardSh.getRange(2, 1, lastInwardRow - 1, 6).getValues();
      inwardData.forEach(row => {
        const itemName = String(row[2] || "").trim().toLowerCase(); 
        const itemDesc = String(row[3] || "").trim().toLowerCase();
        const dId = String(row[5] || "").trim();      
        
        if (dId === String(orderId).trim()) {
          // Unique key using both Item and Description
          const key = itemName + "|" + itemDesc;
          packedCounts[key] = (packedCounts[key] || 0) + 1;
        }
      });
    }
  }

  const data = orderSh.getRange(2, 1, lastOrderRow - 1, 8).getValues(); 
  
  let payload = {
    found: false,
    alreadySubmitted: false, 
    date: "",
    customerName: "", 
    shipTo: "",       
    items: []
  };

  let orderIdExists = false;

  // 2. Iterate through ALL rows in Order sheet
  data.forEach(row => {
    if (String(row[0]).trim() === String(orderId).trim()) {
      orderIdExists = true;
      payload.found = true;
      
      if (!payload.date) {
         let d = row[1];
         try { 
            if (d instanceof Date) { d = d.toISOString().split('T')[0]; }
         } catch(e) { d = row[1]; }
         
         payload.date = d;
         payload.customerName = row[2];
         payload.shipTo = row[3];
      }
      
      const itemNameOrig = String(row[4] || "").trim();
      const itemDescOrig = String(row[5] || "").trim();
      
      if(itemNameOrig) {
        const totalQty = Number(row[6]) || 0;
        
        // Generate Case-Insensitive key for matching
        const matchKey = itemNameOrig.toLowerCase() + "|" + itemDescOrig.toLowerCase();
        
        let currentPacked = packedCounts[matchKey] || 0;
        const deduction = Math.min(totalQty, currentPacked);
        const pendingQty = totalQty - deduction;

        // Reduce currentPacked count so multiple lines of same item are handled correctly
        packedCounts[matchKey] = currentPacked - deduction;

        // Only add to list if there is quantity pending
        if (pendingQty > 0) {
          payload.items.push({
            itemName: itemNameOrig,                  
            itemDescription: itemDescOrig,       
            totalQty: totalQty,                  
            pendingQty: pendingQty,              
            type: row[7] || "Bora"               
          });
        }
      }
    }
  });

  if (orderIdExists && payload.items.length === 0) {
    payload.alreadySubmitted = true;
  } else {
    payload.alreadySubmitted = false;
  }

  return payload;
}

/* ===========================
    DISPATCH ID GENERATOR & PDF LOGIC
=========================== */
function getLastDispatchSeq_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const col = sheet.getRange(2, 6, last - 1, 1).getValues().flat();
  for (let i = col.length - 1; i >= 0; i--) {
    const m = String(col[i]).match(/(\d+)$/);
    if (m) return Number(m[1]);
  }
  return 0;
}

function fmtDispatchId_(prefix, num) {
  return prefix + String(num).padStart(4, "0");
}

function _getChallanHTML_(payload) {
  const {
    dispatchId, date, customerName, shipTo,
    driverName, vehicleNumber, driverNumber,
    loaderName, pdfGenTime,
    items, grandTotalBora, grandTotalWeight
  } = payload;

  const esc = s => String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;");

  const itemRows = items.map(it => {
    let total = it.weights.reduce((a,b)=>a+(+b||0),0);
    let count = it.weights.length;
    let packLabel = it.packingType || "Bora"; 
    let sr = 1;

    const rows = it.weights.map(w => {
      let fullDisplayName = esc(it.itemName);
      if (it.itemDescription) {
          fullDisplayName += " " + esc(it.itemDescription);
      }

      return `
      <div class="row item-row">
        <div class="left">
          ${sr++}. ${fullDisplayName}
        </div>
        <div class="right">${fmt2(w)} kg</div>
      </div>`;
    }).join("");

    let contractorHtml = "";
    if (it.contractorName && String(it.contractorName).trim() !== "") {
        contractorHtml = `
        <div class="subtotal" style="justify-content: flex-end;">
          <span>Contr: <b>${esc(it.contractorName)}</b></span>
        </div>`;
    }

    return `
      <div class="item-block">
        ${rows}
        <div class="subtotal">
          <span>Total Weight: <b>${fmt2(total)}</b> kg</span>
          <span>${esc(packLabel)}: <b>${count}</b></span>
        </div>
        ${contractorHtml}
      </div>`;
  }).join("");

  return `
  <html><head>
  <style>
    body{font-family:Arial;font-size:13px;}
    h2{text-align:center;font-size:24px;margin:0;}
    .hdr{border-bottom:2px solid #000;padding-bottom:10px;margin-bottom:14px;}
    .meta{font-size:13px;margin:2px 0;}
    .row{display:flex;justify-content:space-between;}
    .item-row{font-size:12px;padding:1px 0;}
    .item-block{border-bottom:1px dashed #888;padding:6px 0;margin-bottom:4px;}
    .subtotal{font-size:12px;display:flex;justify-content:space-between;margin-top:2px;}
    .grand{margin-top:12px;padding-top:10px;border-top:2px solid #000;font-size:15px;font-weight:bold;}
  </style>
  </head><body>
    <div class="hdr">
      <h2>Dispatch Challan</h2>
      <div class="row">
        <div class="meta"><b>Date:</b> ${esc(formatPDFDate_(date))}</div>
        <div class="meta"><b>Gen Time:</b> ${esc(pdfGenTime)}</div>
      </div>
      <div class="row">
        <div class="meta"><b>Bill To:</b> ${esc(customerName)}</div>
        <div class="meta"><b>Ship To:</b> ${esc(shipTo)}</div>
      </div>
      <div class="row">
        <div class="meta"><b>Driver:</b> ${esc(driverName)} (${esc(driverNumber)})</div>
        <div class="meta"><b>Vehicle:</b> ${esc(vehicleNumber)}</div>
      </div>
      <div class="row">
        <div class="meta"><b>Loader:</b> ${esc(loaderName)}</div>
        <div class="meta"><b>Dispatch ID:</b> ${esc(dispatchId)}</div>
      </div>
    </div>
    ${itemRows}
    <div class="grand">
      <div>Total Units (Bora/Bale): ${grandTotalBora}</div>
      <div>Grand Total Weight: ${fmt2(grandTotalWeight)} kg</div>
    </div>
  </body></html>`;
}

function _htmlToPdfDataUrl_(html, filename) {
  const blob = Utilities.newBlob(html, "text/html", "challan.html").getAs("application/pdf").setName(filename);
  return "data:application/pdf;base64," + Utilities.base64Encode(blob.getBytes());
}

function _saveFormAndGetPayload_(formData) {
  const { entries, timestamp, duration, driverName, vehicleNumber, driverNumber, loaderName, pdfGenTime, orderId, dispatchId } = formData;
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("INWARD");
  if (!sheet) throw new Error("INWARD sheet not found");

  const PREFIX = "DSP-";
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error("Could not acquire lock");

  try {
    let newDispatchId;
    let incomingId = orderId || dispatchId; 

    if (incomingId && 
        String(incomingId).trim() !== "" && 
        String(incomingId).toLowerCase() !== "undefined" && 
        String(incomingId).toLowerCase() !== "null") {
      newDispatchId = String(incomingId).trim();
    } else {
      const lastSeq = getLastDispatchSeq_(sheet);
      newDispatchId = fmtDispatchId_(PREFIX, lastSeq + 1);
    }

    let headerDate = "", headerCust = "", headerShip = "";
    let sr = sheet.getLastRow();
    const pdfItems = [];

    entries.forEach(({itemName, itemDescription, date, customerName, shipTo, contractorName, packerName, packingType, orderQty, boras = []}) => {
      if (!Array.isArray(boras)) boras = [];
      if (!headerDate) headerDate = date;
      if (!headerCust) headerCust = customerName;
      if (!headerShip) headerShip = shipTo;
      const pType = packingType || "Bora";

      boras.forEach((w, i) => {
        sr++;
        sheet.appendRow([
          sr, 
          timestamp, 
          itemName, 
          itemDescription || "", 
          date, 
          newDispatchId, 
          customerName, 
          shipTo, 
          i + 1, 
          w, 
          contractorName, 
          packerName || "", 
          duration + " sec", 
          driverName, 
          vehicleNumber, 
          driverNumber, 
          loaderName, 
          pdfGenTime, 
          pType,
          orderQty || "" 
        ]);
      });
      pdfItems.push({ itemName, itemDescription, weights: boras, contractorName, packingType: pType });
    });

    return {
      dispatchId: newDispatchId, date: headerDate, customerName: headerCust, shipTo: headerShip,
      driverName, vehicleNumber, driverNumber, loaderName, pdfGenTime, items: pdfItems,
      grandTotalBora: entries.reduce((t, e) => t + e.boras.length, 0),
      grandTotalWeight: entries.reduce((t, e) => t + e.boras.reduce((x, y) => x + (+y || 0), 0), 0)
    };
  } finally { lock.releaseLock(); }
}

function submitFormAndCreatePDF(formData) {
  try {
    const p = _saveFormAndGetPayload_(formData);
    const pForPdf = Object.assign({}, p, { pdfGenTime: formatPDFTime_(p.pdfGenTime) });
    const htmlPdf = _getChallanHTML_(pForPdf);
    const file = p.customerName + "_" + p.dispatchId + ".pdf";
    let url = _htmlToPdfDataUrl_(htmlPdf, file);
    if (!_isValidPdfDataUrl_(url)) url = _driveHtmlToPdfDataUrl_(htmlPdf, file);
    return { success:true, dataUrl: url, filename: file };
  } catch (err) { return { success:false, error: err.message || String(err) }; }
}

function _isValidPdfDataUrl_(d) { try{ return Utilities.base64Decode(d.split(",")[1]).length > 10; }catch(e){return false;} }
function _driveHtmlToPdfDataUrl_(h, f) {
  const t = DriveApp.createFile("temp.html", h, MimeType.HTML);
  const b = t.getAs(MimeType.PDF).setName(f).getBytes();
  DriveApp.removeFile(t);
  return "data:application/pdf;base64," + Utilities.base64Encode(b);
}

/* ===========================
    FIXED REPRINT LOGIC
    Groups by Item + Description + Contractor separately
=========================== */
function createChallanPDF(dispatchIdInput) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("INWARD");
  if (!sh) throw new Error("INWARD sheet not found");

  const last = sh.getLastRow();
  if (last < 2) throw new Error("No data found in INWARD sheet");

  const data = sh.getRange(2, 1, last - 1, 20).getValues();
  
  let date="", cust="", ship="", driver="", veh="", dnum="", loader="", gtime="";
  const map = {};

  data.forEach(r => {
    // Array Mapping based on spreadsheet structure
    const [sr, ts, item, desc, dt, dId, cu, sp, no, w, cont, pack, dur, dN, vN, dM, lN, pT, pType, oQty] = r;
    
    if (String(dId).trim() !== String(dispatchIdInput).trim()) return;

    if (!date) {
      date = dt; cust = cu; ship = sp;
      driver = dN; veh = vN; dnum = dM; loader = lN; gtime = pT;
    }

    // UPDATE: groupKey now uses Item Name + Description + Contractor 
    // This ensures they show up as separate blocks if any of these 3 are different
    const groupKey = String(item) + "|" + String(desc || "") + "|" + String(cont || "");
    
    if (!map[groupKey]) {
      map[groupKey] = { 
        itemName: item,
        weights: [], 
        itemDescription: desc, 
        contractor: cont, 
        packingType: pType || "Bora" 
      };
    }
    if (w) map[groupKey].weights.push(w);
  });

  if (!date) throw new Error("Challan Number " + dispatchIdInput + " not found in INWARD records.");

  const items = Object.keys(map).map(k => ({
    itemName: map[k].itemName, 
    itemDescription: map[k].itemDescription, 
    weights: map[k].weights,
    contractorName: map[k].contractor, 
    packingType: map[k].packingType
  }));

  const payload = {
    dispatchId: dispatchIdInput, 
    date: date, 
    customerName: cust, 
    shipTo: ship,
    driverName: driver, 
    vehicleNumber: veh, 
    driverNumber: dnum,
    loaderName: loader, 
    pdfGenTime: gtime, 
    items: items,
    grandTotalBora: items.reduce((t,it)=>t+it.weights.length, 0),
    grandTotalWeight: items.reduce((t,it)=>t+it.weights.reduce((a,b)=>a+(+b||0),0),0)
  };
  
  const payloadForPdf = Object.assign({}, payload, { pdfGenTime: formatPDFTime_(payload.pdfGenTime) });
  const htmlPdf = _getChallanHTML_(payloadForPdf);
  const file = cust + "_" + dispatchIdInput + ".pdf";
  
  let url = _htmlToPdfDataUrl_(htmlPdf, file);
  if (!_isValidPdfDataUrl_(url)) url = _driveHtmlToPdfDataUrl_(htmlPdf, file);
  
  return { dataUrl:url, filename:file };
}