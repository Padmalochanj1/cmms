// ===== Cell map for CMMS_CWEI_V1_3.xlsx - change here if the layout moves =====
const CFG = {
  closing: {
    sheet: "Closing Maintenance Orders",
    source: "B9:M308",          // B=Order, C=EQ ID, D=Name, E=Type, F=Location, G=Status, H=Action, I=Closed On, J=Closed By, K=Remarks, L=Page, M=Checklist Status
    message: "A6",
    calcSheet: "Calc",
    calc: "AK2:AR1000",         // AK=Order ... AQ=Frequency, AR=Due Date
    table: "MaintenanceClosureHistory",
    columns: 9
  },
  checklist: {
    sheet: "Order Checklist",
    order: "B3",
    status: "H5",
    export: "A37:T56",          // 20 rows x 20 columns, same order as MeasurementHistory
    clear: ["F10:F29", "I10:J29"], // Actual Reading, then OK / NOT OK + Notes
    message: "A7",
    table: "MeasurementHistory",
    columns: 20
  }
};

Office.onReady(() => {
  document.getElementById("saveClosure").onclick = () => run("saveClosure", saveClosures);
  document.getElementById("saveChecklist").onclick = () => run("saveChecklist", saveChecklist);
  watchOrderChange();
});

// Clears the checklist entries whenever a different order is selected in B3.
// Works while the CMMS Save pane is open.
async function watchOrderChange() {
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(CFG.checklist.sheet);
      ws.onChanged.add(onChecklistSheetChanged);
      await context.sync();
    });
  } catch (e) { /* the save buttons still work without the watcher */ }
}

async function onChecklistSheetChanged(ev) {
  const K = CFG.checklist;
  const addr = String(ev.address || "").split("!").pop().replace(/\$/g, "");
  if (addr.split(":")[0] !== K.order) return;   // only react to the order cell (B3 / merged B3:C3)
  try {
    await Excel.run(async (context) => {
      const sh = context.workbook.worksheets.getItem(K.sheet);
      const st = await openSheet(context, sh, K.sheet);
      try {
        K.clear.forEach(a => sh.getRange(a).clear(Excel.ClearApplyTo.contents));
        sh.getRange(K.message).clear(Excel.ClearApplyTo.contents);
        await context.sync();
      } finally {
        await closeSheets(context, [st]);
      }
    });
    showMessage("Order changed - checklist entries cleared. Fill in the new checklist.");
  } catch (e) {
    showMessage("Could not clear the previous entries: " + (e.message || e), true);
  }
}

function showMessage(text, isError = false) {
  const el = document.getElementById("message");
  el.textContent = text;
  el.style.borderColor = isError ? "#c50f1f" : "#107c10";
}

// Popup over Excel. type: "ok" | "warn" | "err"
function notify(type, msg) {
  try {
    const url = new URL("dialog.html", window.location.href);
    url.searchParams.set("type", type);
    url.searchParams.set("msg", String(msg).slice(0, 300));
    Office.context.ui.displayDialogAsync(
      url.toString(),
      { height: 28, width: 28, displayInIframe: true },
      (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) return;
        const dlg = res.value;
        dlg.addEventHandler(Office.EventType.DialogMessageReceived, () => dlg.close());
      }
    );
  } catch (e) { /* the pane message is the fallback */ }
}

async function run(buttonId, fn) {
  const btn = document.getElementById(buttonId);
  btn.disabled = true;
  showMessage("Working...");
  try {
    const text = await fn();
    const saved = text.startsWith("Saved");
    showMessage(text, !saved);
    notify(saved ? "ok" : "warn", text.split("\n")[0]);
  } catch (error) {
    const m = error.message || String(error);
    showMessage("ERROR: " + m, true);
    notify("err", m);
  } finally {
    btn.disabled = false;
  }
}

const txt = v => String(v === null || v === undefined ? "" : v).trim();

function stampNow() {
  return new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });
}

// Confirms sheets and table exist before doing anything else.
async function checkItems(context, sheetNames, tableName) {
  const sheets = sheetNames.map(n => {
    const s = context.workbook.worksheets.getItemOrNullObject(n);
    s.load("isNullObject");
    return { n, s };
  });
  const tbl = context.workbook.tables.getItemOrNullObject(tableName);
  tbl.load("isNullObject");
  await context.sync();
  for (const x of sheets) {
    if (x.s.isNullObject) throw new Error(`Sheet "${x.n}" was not found. Check the sheet name.`);
  }
  if (tbl.isNullObject) throw new Error(`Table "${tableName}" was not found. Check the table name (Table Design tab).`);
  return tbl;
}

// Temporarily unprotects a sheet (only works when the sheet has NO password).
async function openSheet(context, sheet, name) {
  sheet.protection.load("protected,options");
  await context.sync();
  const st = { sheet, name, was: sheet.protection.protected, options: sheet.protection.options };
  if (st.was) {
    try {
      sheet.protection.unprotect();
      await context.sync();
    } catch (e) {
      throw new Error(`Sheet "${name}" is password-protected, so the add-in cannot write to it. Remove its protection once (Review > Unprotect Sheet) and try again.`);
    }
  }
  return st;
}

async function closeSheets(context, opened) {
  for (const st of opened) {
    if (!st.was) continue;
    try {
      st.sheet.protection.protect(st.options);
      await context.sync();
    } catch (e) { /* ignore: leave unprotected rather than fail the save */ }
  }
}

async function saveClosures() {
  const C = CFG.closing;
  return Excel.run(async (context) => {
    const tbl = await checkItems(context, [C.sheet, C.calcSheet], C.table);
    const sh = context.workbook.worksheets.getItem(C.sheet);
    const source = sh.getRange(C.source);
    const calc = context.workbook.worksheets.getItem(C.calcSheet).getRange(C.calc);
    const tableRange = tbl.getRange();
    source.load("values");
    calc.load("values");
    tableRange.load("values");
    await context.sync();

    const colCount = tableRange.values[0].length;
    if (colCount !== C.columns) {
      throw new Error(`${C.table} has ${colCount} columns but ${C.columns} are needed.`);
    }

    const details = new Map();
    for (const r of calc.values) {
      if (txt(r[0])) details.set(txt(r[0]), r);
    }
    const existing = new Set(
      tableRange.values.slice(1).filter(r => txt(r[0])).map(r => txt(r[0]))
    );

    const withOrder = source.values.filter(r => txt(r[0]) !== "");
    const deferred = withOrder.filter(r => txt(r[5]).toUpperCase() === "DEFERRED");
    const closed = withOrder.filter(r => txt(r[5]).toUpperCase() === "CLOSED");
    const ready = closed.filter(r => txt(r[11]).toUpperCase() === "COMPLETE");

    const add = [];
    const missing = [];
    let alreadySaved = 0;
    for (const r of ready) {
      const orderNo = txt(r[0]);
      if (existing.has(orderNo)) { alreadySaved++; continue; }
      const d = details.get(orderNo);
      if (!d || txt(d[7]) === "") { missing.push(orderNo); continue; }
      // Order, Due Date, EQ ID, Equipment Name, Frequency, Action, Closed On, Closed By, Remarks
      add.push([orderNo, d[7], r[1], r[2], d[6], r[6], r[7], r[8], r[9]]);
    }

    const report =
      `Orders listed: ${withOrder.length}\n` +
      `Status CLOSED: ${closed.length}  (Deferred, not saved: ${deferred.length})\n` +
      `Checklist COMPLETE: ${ready.length}\n` +
      `Already in history: ${alreadySaved}\n` +
      `Due date not found: ${missing.length}${missing.length ? " (" + missing.join(", ") + ")" : ""}\n` +
      `To save now: ${add.length}`;

    if (!add.length) return "Nothing saved.\n" + report;

    const opened = [];
    try {
      opened.push(await openSheet(context, tbl.worksheet, "Maintenance Closure History"));
      opened.push(await openSheet(context, sh, C.sheet));
      tbl.rows.add(null, add);
      sh.getRange(C.message).values = [[`SAVED - ${add.length} order(s) to Maintenance Closure History on ${stampNow()}.`]];
      const count = tbl.rows.getCount();
      await context.sync();
      return `Saved ${add.length} order(s). History now has ${count.value} rows.\n` + report;
    } finally {
      await closeSheets(context, opened);
    }
  });
}

async function saveChecklist() {
  const K = CFG.checklist;
  return Excel.run(async (context) => {
    const tbl = await checkItems(context, [K.sheet], K.table);
    const sh = context.workbook.worksheets.getItem(K.sheet);
    const orderCell = sh.getRange(K.order);
    const statusCell = sh.getRange(K.status);
    const exportRange = sh.getRange(K.export);
    const tableRange = tbl.getRange();
    orderCell.load("values");
    statusCell.load("values");
    exportRange.load("values");
    tableRange.load("values");
    await context.sync();

    const colCount = tableRange.values[0].length;
    if (colCount !== K.columns) {
      throw new Error(`${K.table} has ${colCount} columns but ${K.columns} are needed (same as ${K.export}).`);
    }

    const orderNo = txt(orderCell.values[0][0]);
    const status = txt(statusCell.values[0][0]).toUpperCase();
    if (!orderNo) throw new Error(`Select a maintenance order in Order Checklist cell ${K.order}.`);
    if (status !== "COMPLETE") throw new Error(`Checklist is not complete (${K.status} shows "${status || "blank"}"). Finish every item and enter Recorded By.`);

    const forOrder = exportRange.values.filter(r => txt(r[0]) === orderNo);
    const withItem = forOrder.filter(r => txt(r[8]) !== "");
    const withResult = withItem.filter(r => ["PASS", "ATTENTION"].includes(txt(r[16]).toUpperCase()));
    const rows = withResult.filter(r => txt(r[18]) !== "" && txt(r[19]) !== "");

    const existing = new Set(
      tableRange.values.slice(1).filter(r => txt(r[0]) && txt(r[8]))
        .map(r => txt(r[0]) + "|" + txt(r[8]))
    );
    const add = rows
      .filter(r => !existing.has(txt(r[0]) + "|" + txt(r[8])))
      .map(r => { const x = r.slice(); if (x[17] === 0) x[17] = ""; return x; }); // blank Notes come through as 0

    const report =
      `Rows for order ${orderNo} in ${K.export}: ${forOrder.length}\n` +
      `With item filled (column I): ${withItem.length}\n` +
      `Result PASS/ATTENTION (column Q): ${withResult.length}\n` +
      `Recorded By + date filled (columns S, T): ${rows.length}\n` +
      `Already in history: ${rows.length - add.length}\n` +
      `To save now: ${add.length}`;

    if (!add.length) return "Nothing saved.\n" + report;

    const opened = [];
    try {
      opened.push(await openSheet(context, tbl.worksheet, "Measurement History"));
      opened.push(await openSheet(context, sh, K.sheet));
      tbl.rows.add(null, add);
      K.clear.forEach(a => sh.getRange(a).clear(Excel.ClearApplyTo.contents));
      sh.getRange(K.message).values = [[`SAVED - ${orderNo} (${add.length} items) to Measurement History on ${stampNow()}. Entries cleared.`]];
      const count = tbl.rows.getCount();
      await context.sync();
      return `Saved ${add.length} item(s). History now has ${count.value} rows.\n` + report;
    } finally {
      await closeSheets(context, opened);
    }
  });
}
