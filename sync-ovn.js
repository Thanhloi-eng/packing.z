const crypto = require("crypto");
const XLSX = require("xlsx");
// Single source of truth for Sheet "OVN SCHEDULE" (146 columns from Ovn Pro Schedule.xlsb)
// Keep synchronized with supabase_ovn_schema.sql and api/sync-ovn.js

const EXACT_HEADERS = [
  "So", "PRO ORDER", "Brand", "Customer", "Type Oder", "#MOLDTYPE", "QtyOrder",
  "loadMaterial(PPC)", "Recieved Material", "Recieved Logo", "Lamination(PPC)", "LAMINATION(PRO)",
  "Sawcutting(PPC)", "PRE(PRO)", "Tách Bào(PPC)", "Slipting(PRO)", "SUB(PPC)", "Sub Return",
  "SubIFM", "Instruction Sub", "MOLD_IN(PRO)", "MOLDING(PPC)", "MOLD_OUT(PRO)", "LEAN_IN(PRO)",
  "INLEANLINE(PPC)", "OUTLEANLINE(PPC)", "LEAN_OUT(PRO)", "LINE CODE", "Returned Line", "KHO _TẠM",
  "STORED", "Finish Date(PPC)", "PPC CMF", "Status", "BOM", "#LAST", "Instruction Skiving",
  "ArticleCode", "Gender", "CODE PU1", "Description PU1", "DL PU1", "CODE PU2", "Description PU2",
  "DL PU2", "CODE PU3", "Description PU3", "DL PU3", "CODE FABRIC", "Description FB", "DL FB",
  "CODE LOGO1", "Description LOGO1", "DL LOGO1", "CODE LOGO2", "Description LOGO2", "DL LOGO2",
  "CODE LOGO3", "Description LOGO3", "DL LOGO3", "CODE LOGO4", "Description LOGO4", "DL LOGO4",
  "S1", "S1_5", "S2", "S2.5", "S3", "S3.5", "S4", "S4.5", "S5", "S5.5", "S6", "S6.5", "S7", "S7.5",
  "S8", "S8.5", "S9", "S9.5", "S10", "S10.5", "S11", "S11.5", "S12", "S12.5", "S13", "S13.5",
  "S14", "S14.5", "S15", "S15.5", "S16", "S16.5", "S17", "S18", "S19", "S20", "S21", "S22",
  "S23", "S24", "S25", "S26", "S27", "S28", "S29", "S30", "S31", "S32", "S33", "S34", "S35",
  "S36", "S37", "S38", "S39", "S40", "S41", "S42", "S43", "S44", "S45", "S46", "S47", "S48",
  "S49", "S3.5Y", "S4Y", "S4.5Y", "S5Y", "S5.5Y", "S6Y", "S6.5Y", "S7Y", "S10K", "S10.5K",
  "S11K", "S11.5K", "S12K", "NG Fabric", "Inventory_Logo_Inhouse", "NOTE", "SUPP", "Qty_SUPP"
];

const SPECIAL_SQL_NAMES = {
  "So": "so",
  "PRO ORDER": "pro_order",
  "Brand": "brand",
  "Customer": "customer",
  "Type Oder": "type_oder",
  "#MOLDTYPE": "mold_type",
  "QtyOrder": "qty_order",
  "loadMaterial(PPC)": "load_material_ppc",
  "Recieved Material": "recieved_material",
  "Recieved Logo": "recieved_logo",
  "Lamination(PPC)": "lamination_ppc",
  "LAMINATION(PRO)": "lamination_pro",
  "Sawcutting(PPC)": "sawcutting_ppc",
  "PRE(PRO)": "pre_pro",
  "Tách Bào(PPC)": "tach_bao_ppc",
  "Slipting(PRO)": "slipting_pro",
  "SUB(PPC)": "sub_ppc",
  "Sub Return": "sub_return",
  "SubIFM": "sub_ifm",
  "Instruction Sub": "instruction_sub",
  "MOLD_IN(PRO)": "mold_in_pro",
  "MOLDING(PPC)": "molding_ppc",
  "MOLD_OUT(PRO)": "mold_out_pro",
  "LEAN_IN(PRO)": "lean_in_pro",
  "INLEANLINE(PPC)": "inleanline_ppc",
  "OUTLEANLINE(PPC)": "outleanline_ppc",
  "LEAN_OUT(PRO)": "lean_out_pro",
  "LINE CODE": "line_code",
  "Returned Line": "returned_line",
  "KHO _TẠM": "kho_tam",
  "STORED": "stored",
  "Finish Date(PPC)": "finish_date_ppc",
  "PPC CMF": "ppc_cmf",
  "Status": "status",
  "BOM": "bom",
  "#LAST": "last_code",
  "Instruction Skiving": "instruction_skiving",
  "ArticleCode": "article_code",
  "Gender": "gender",
  "CODE PU1": "code_pu1",
  "Description PU1": "description_pu1",
  "DL PU1": "dl_pu1",
  "CODE PU2": "code_pu2",
  "Description PU2": "description_pu2",
  "DL PU2": "dl_pu2",
  "CODE PU3": "code_pu3",
  "Description PU3": "description_pu3",
  "DL PU3": "dl_pu3",
  "CODE FABRIC": "code_fabric",
  "Description FB": "description_fb",
  "DL FB": "dl_fb",
  "CODE LOGO1": "code_logo1",
  "Description LOGO1": "description_logo1",
  "DL LOGO1": "dl_logo1",
  "CODE LOGO2": "code_logo2",
  "Description LOGO2": "description_logo2",
  "DL LOGO2": "dl_logo2",
  "CODE LOGO3": "code_logo3",
  "Description LOGO3": "description_logo3",
  "DL LOGO3": "dl_logo3",
  "CODE LOGO4": "code_logo4",
  "Description LOGO4": "description_logo4",
  "DL LOGO4": "dl_logo4",
  "NG Fabric": "ng_fabric",
  "Inventory_Logo_Inhouse": "inventory_logo_inhouse",
  "NOTE": "note",
  "SUPP": "supp",
  "Qty_SUPP": "qty_supp"
};

const TIMESTAMP_EXCEL_COLS = new Set([
  "LAMINATION(PRO)", "PRE(PRO)", "Slipting(PRO)",
  "MOLD_IN(PRO)", "MOLD_OUT(PRO)", "LEAN_IN(PRO)", "LEAN_OUT(PRO)"
]);

const DATE_EXCEL_COLS = new Set([
  "loadMaterial(PPC)", "Recieved Material", "Recieved Logo", "Lamination(PPC)",
  "Sawcutting(PPC)", "Tách Bào(PPC)", "SUB(PPC)", "Sub Return", "MOLDING(PPC)",
  "INLEANLINE(PPC)", "OUTLEANLINE(PPC)", "Returned Line", "KHO _TẠM", "STORED",
  "Finish Date(PPC)", "PPC CMF"
]);

const NUMERIC_EXCEL_COLS = new Set([
  "QtyOrder", "DL PU1", "DL PU2", "DL PU3", "DL FB",
  "DL LOGO1", "DL LOGO2", "DL LOGO3", "DL LOGO4",
  "NG Fabric", "Inventory_Logo_Inhouse", "Qty_SUPP"
]);

function toSqlName(excelHeader) {
  if (SPECIAL_SQL_NAMES[excelHeader]) return SPECIAL_SQL_NAMES[excelHeader];
  return String(excelHeader)
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d").replace(/Đ/g, "D")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function getSqlType(excelHeader, idx) {
  if (TIMESTAMP_EXCEL_COLS.has(excelHeader)) return "TIMESTAMPTZ";
  if (DATE_EXCEL_COLS.has(excelHeader)) return "DATE";
  if (NUMERIC_EXCEL_COLS.has(excelHeader) || (idx >= 63 && idx <= 140)) return "NUMERIC";
  return "TEXT";
}

const OVN_COLUMNS = EXACT_HEADERS.map((excelName, idx) => ({
  idx,
  excelName,
  sqlName: toSqlName(excelName),
  sqlType: getSqlType(excelName, idx)
}));

const DEFAULT_SUPABASE_URL = "https://plimlhiopeloywyfxdgk.supabase.co";
const DEFAULT_DRIVE_FILE_ID = "125H4VKxN4ezLDQD_JB4ETeFtmRCxm1aq";
const SHEET_PASSWORD_HINT = process.env.OVN_SHEET_PASSWORD || "MOLDING@2024";
const BATCH_SIZE = 500;
const CONCURRENCY = 1;

const pad2 = (n) => String(n).padStart(2, "0");

function isZipBinary(buf) {
  return buf && buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b;
}

async function downloadDriveFile(fileId) {
  const initialUrl = `https://drive.google.com/uc?export=download&id=${encodeURIComponent(fileId)}`;
  let resp = await fetch(initialUrl, { redirect: "follow" });
  if (!resp.ok) {
    throw new Error(`Google Drive HTTP ${resp.status} while downloading ${fileId}`);
  }

  let buf = Buffer.from(await resp.arrayBuffer());
  if (isZipBinary(buf)) return buf;

  // Handle Google Drive virus-scan interstitial page for larger files
  const html = buf.toString("utf8");
  const uuidMatch = html.match(/name="uuid"\s+value="([^"]+)"/i);
  const confirmMatch = html.match(/name="confirm"\s+value="([^"]+)"/i) || html.match(/confirm=([0-9A-Za-z_-]+)/);
  const confirmToken = confirmMatch ? confirmMatch[1] : "t";
  let confirmUrl = `https://drive.usercontent.google.com/download?id=${encodeURIComponent(fileId)}&export=download&confirm=${encodeURIComponent(confirmToken)}`;
  if (uuidMatch) {
    confirmUrl += `&uuid=${encodeURIComponent(uuidMatch[1])}`;
  }

  resp = await fetch(confirmUrl, { redirect: "follow" });
  if (!resp.ok) {
    throw new Error(`Google Drive confirm download HTTP ${resp.status}`);
  }
  buf = Buffer.from(await resp.arrayBuffer());
  if (!isZipBinary(buf)) {
    throw new Error("Downloaded payload is not a valid .xlsb/.xlsx file (check Drive public sharing permission).");
  }
  return buf;
}

function excelSerialToParts(val) {
  if (val === null || val === undefined || val === "") return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    const serial = 25569 + val.getTime() / 86400000;
    return {
      serial,
      y: val.getUTCFullYear(),
      m: val.getUTCMonth() + 1,
      d: val.getUTCDate(),
      hh: val.getUTCHours(),
      mm: val.getUTCMinutes(),
      ss: val.getUTCSeconds(),
      hasTime: val.getUTCHours() > 0 || val.getUTCMinutes() > 0 || val.getUTCSeconds() > 0
    };
  }
  const num = Number(val);
  if (!Number.isNaN(num) && num >= 35000 && num <= 60000) {
    const dt = new Date(Math.round((num - 25569) * 86400000));
    const frac = num - Math.floor(num);
    return {
      serial: num,
      y: dt.getUTCFullYear(),
      m: dt.getUTCMonth() + 1,
      d: dt.getUTCDate(),
      hh: dt.getUTCHours(),
      mm: dt.getUTCMinutes(),
      ss: dt.getUTCSeconds(),
      hasTime: frac > 0.0001
    };
  }
  if (typeof val === "string") {
    const s = val.trim();
    const m =
      s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/) ||
      s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (m) {
      const isIso = m[1].length === 4;
      const y = Number(isIso ? m[1] : m[3]);
      const mo = Number(m[2]);
      const d = Number(isIso ? m[3] : m[1]);
      const hh = Number(m[4] || 0);
      const mm = Number(m[5] || 0);
      const ss = Number(m[6] || 0);
      const utcMs = Date.UTC(y, mo - 1, d, hh, mm, ss);
      return {
        serial: 25569 + utcMs / 86400000,
        y,
        m: mo,
        d,
        hh,
        mm,
        ss,
        hasTime: Boolean(m[4])
      };
    }
  }
  return null;
}

function formatSqlValue(rawVal, sqlType) {
  if (rawVal === null || rawVal === undefined || rawVal === "") return null;

  if (sqlType === "TIMESTAMPTZ") {
    const p = excelSerialToParts(rawVal);
    if (!p) return null;
    return `${p.y}-${pad2(p.m)}-${pad2(p.d)}T${pad2(p.hh)}:${pad2(p.mm)}:${pad2(p.ss)}+07:00`;
  }

  if (sqlType === "DATE") {
    const p = excelSerialToParts(rawVal);
    if (!p) return null;
    return `${p.y}-${pad2(p.m)}-${pad2(p.d)}`;
  }

  if (sqlType === "NUMERIC") {
    if (typeof rawVal === "number") return Number.isFinite(rawVal) ? rawVal : null;
    const cleaned = String(rawVal).replace(/,/g, "").trim();
    if (!cleaned) return null;
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : null;
  }

  const s = String(rawVal).trim();
  return s.length > 0 ? s : null;
}

function findOvnSheet(workbook) {
  const names = workbook.SheetNames || [];
  const exact = names.find((n) => /^ovn\s*schedule$/i.test(n.trim()));
  if (exact) return exact;
  const partial = names.find((n) => /^ovn/i.test(n.trim()));
  if (partial) return partial;
  return names[0];
}

function parseOvnWorkbook(buffer) {
  const wb = XLSX.read(buffer, {
    type: "buffer",
    raw: true,
    cellDates: false,
    password: SHEET_PASSWORD_HINT
  });
  const sheetName = findOvnSheet(wb);
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet "${sheetName}" not found in workbook`);

  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
  if (!matrix || matrix.length < 2) {
    throw new Error(`Sheet "${sheetName}" contains no data rows`);
  }

  const excelHeaders = matrix[0].map((h) => String(h || "").trim());
  const colIndexBySql = new Map();
  const knownBySql = new Map(OVN_COLUMNS.map((c) => [c.sqlName, c]));

  for (let c = 0; c < excelHeaders.length; c++) {
    const h = excelHeaders[c];
    if (!h) continue;
    const sqlName = toSqlName(h);
    if (knownBySql.has(sqlName) && !colIndexBySql.has(sqlName)) {
      colIndexBySql.set(sqlName, c);
    }
  }

  if (!colIndexBySql.has("pro_order")) {
    throw new Error(`Required column "PRO ORDER" missing in sheet "${sheetName}"`);
  }

  const proColIdx = colIndexBySql.get("pro_order");
  const typeColIdx = colIndexBySql.get("type_oder");
  const lineColIdx = colIndexBySql.get("line_code");

  const sqlRowsMap = new Map();
  const diecutRawRows = [];
  const stations = {};
  let latestFloorScan = null;
  let latestPlanDate = null;
  let totalScans = 0;

  const isFloorStation = (h) =>
    (/\(pro\)|lean[_\s]*(?:in|out)|mold[_\s]*(?:in|out)|slipting|pre\s*\(pro\)|lamination\s*\(pro\)/i.test(h)) &&
    !/\(ppc\)|finish\s*date|ppc\s*cmf/i.test(h);

  const isPlanCol = (h) => /\(ppc\)|finish\s*date|ppc\s*cmf/i.test(h);

  const nowIso = new Date().toISOString();

  for (let r = 1; r < matrix.length; r++) {
    const rawRow = matrix[r];
    if (!rawRow || rawRow.length === 0) continue;

    const proVal = String(rawRow[proColIdx] || "").trim();
    if (!proVal) continue;

    const sqlObj = { updated_at: nowIso };
    for (const colDef of OVN_COLUMNS) {
      const idx = colIndexBySql.get(colDef.sqlName);
      sqlObj[colDef.sqlName] = idx !== undefined ? formatSqlValue(rawRow[idx], colDef.sqlType) : null;
    }
    sqlRowsMap.set(proVal, sqlObj);

    const lineVal = lineColIdx !== undefined ? String(rawRow[lineColIdx] || "").trim() : "";

    // Compute scan telemetry across all OVN rows
    for (let c = 0; c < excelHeaders.length; c++) {
      const hName = excelHeaders[c];
      if (!hName || (!TIMESTAMP_EXCEL_COLS.has(hName) && !DATE_EXCEL_COLS.has(hName))) continue;
      const parts = excelSerialToParts(rawRow[c]);
      if (!parts) continue;
      totalScans++;

      const formattedTime = parts.hasTime
        ? `${pad2(parts.hh)}:${pad2(parts.mm)}:${pad2(parts.ss)} (${pad2(parts.d)}/${pad2(parts.m)}/${parts.y})`
        : `${pad2(parts.d)}/${pad2(parts.m)}/${parts.y}`;

      if (isFloorStation(hName)) {
        const hit = {
          station: hName,
          proOrder: proVal,
          line: lineVal,
          serial: parts.serial,
          hasTime: parts.hasTime,
          formattedTime
        };
        if (!stations[hName] || parts.serial > stations[hName].serial) {
          stations[hName] = hit;
        }
        if (parts.hasTime && (!latestFloorScan || parts.serial > latestFloorScan.serial)) {
          latestFloorScan = hit;
        }
      } else if (isPlanCol(hName)) {
        if (!latestPlanDate || parts.serial > latestPlanDate.serial) {
          latestPlanDate = {
            station: hName,
            proOrder: proVal,
            line: lineVal,
            serial: parts.serial,
            hasTime: parts.hasTime,
            formattedTime
          };
        }
      }
    }

    // Collect 72-column snapshot for legacy europlan_sync (DIE CUT active focus)
    const typeStr = typeColIdx !== undefined ? String(rawRow[typeColIdx] || "").toLowerCase() : "";
    if (/die\s*cut/.test(typeStr)) {
      const trimmed = new Array(Math.min(72, excelHeaders.length));
      for (let k = 0; k < trimmed.length; k++) {
        trimmed[k] = rawRow[k] !== undefined && rawRow[k] !== null ? String(rawRow[k]) : "";
      }
      diecutRawRows.push(trimmed);
    }
  }

  return {
    sheetName,
    sqlRows: Array.from(sqlRowsMap.values()),
    diecutHeaders: excelHeaders.slice(0, Math.min(72, excelHeaders.length)),
    diecutRawRows,
    telemetry: {
      totalScans,
      latestFloorScan,
      latestPlanDate,
      stations,
      totalOvnRows: sqlRowsMap.size
    }
  };
}

async function supabaseRequest(url, serviceKey, method, bodyObj, preferHeader) {
  const headers = {
    apikey: serviceKey,
    Authorization: `Bearer ${serviceKey}`,
    "Content-Type": "application/json"
  };
  if (preferHeader) headers.Prefer = preferHeader;

  const resp = await fetch(url, {
    method,
    headers,
    body: bodyObj !== undefined ? JSON.stringify(bodyObj) : undefined
  });

  if (!resp.ok) {
    const errText = await resp.text();
    throw new Error(`Supabase ${method} ${resp.status}: ${errText.slice(0, 400)}`);
  }
  const text = await resp.text();
  return text ? JSON.parse(text) : null;
}

async function batchUpsertOvnSchedule(supabaseUrl, serviceKey, rows) {
  const chunks = [];
  for (let i = 0; i < rows.length; i += BATCH_SIZE) {
    chunks.push(rows.slice(i, i + BATCH_SIZE));
  }

  const endpoint = `${supabaseUrl}/rest/v1/ovn_schedule?on_conflict=pro_order`;
  let cursor = 0;

  async function worker() {
    while (cursor < chunks.length) {
      const idx = cursor++;
      await supabaseRequest(
        endpoint,
        serviceKey,
        "POST",
        chunks[idx],
        "resolution=merge-duplicates,return=minimal"
      );
    }
  }

  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, chunks.length) }, worker));
}

async function updateEuroplanSyncSnapshot(supabaseUrl, serviceKey, parsed, existingTelemetry, fileSha256) {
  const physicalInventory =
    (existingTelemetry && existingTelemetry.physicalInventory) ? existingTelemetry.physicalInventory : {};

  const nowIso = new Date().toISOString();
  const payload = {
    id: "latest",
    source_name: "Ovn Pro Schedule.xlsb",
    last_sync: nowIso,
    updated_at: nowIso,
    headers: parsed.diecutHeaders,
    rows: parsed.diecutRawRows,
    telemetry: {
      ...parsed.telemetry,
      fileSha256,
      physicalInventory
    }
  };

  await supabaseRequest(
    `${supabaseUrl}/rest/v1/europlan_sync?on_conflict=id`,
    serviceKey,
    "POST",
    [payload],
    "resolution=merge-duplicates,return=minimal"
  );
}

function isAuthorized(req) {
  const secret = process.env.SYNC_SECRET;
  const cronSecret = process.env.CRON_SECRET;
  if (!secret && !cronSecret) return true;

  const headerSecret = req.headers["x-sync-secret"];
  const authHeader = req.headers["authorization"] || "";
  if (secret && headerSecret === secret) return true;
  if (cronSecret && authHeader === `Bearer ${cronSecret}`) return true;
  if (secret && authHeader === `Bearer ${secret}`) return true;
  return false;
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, x-sync-secret");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (!isAuthorized(req)) return res.status(401).json({ ok: false, error: "Unauthorized" });

  const startedAt = Date.now();
  const supabaseUrl = (process.env.SUPABASE_URL || DEFAULT_SUPABASE_URL).replace(/\/+$/, "");
  const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SECRET_KEY;

  if (!serviceKey) {
    return res.status(500).json({
      ok: false,
      error: "Missing SUPABASE_SERVICE_ROLE_KEY in Vercel Environment Variables."
    });
  }

  try {
    const fileId =
      (req.body && req.body.fileId) ||
      (req.query && req.query.fileId) ||
      process.env.GOOGLE_DRIVE_FILE_ID ||
      DEFAULT_DRIVE_FILE_ID;

    const fileBuffer = await downloadDriveFile(fileId);
    const fileSha256 = crypto.createHash("sha256").update(fileBuffer).digest("hex");
    const forceSync = String((req.query && req.query.force) || (req.body && req.body.force) || "") === "1";

    let existingTelemetry = {};
    try {
      const existing = await supabaseRequest(
        `${supabaseUrl}/rest/v1/europlan_sync?id=eq.latest&select=telemetry`,
        serviceKey,
        "GET"
      );
      if (existing && existing[0] && existing[0].telemetry) {
        existingTelemetry = existing[0].telemetry;
      }
    } catch (_) {}

    if (!forceSync && existingTelemetry.fileSha256 === fileSha256) {
      return res.status(200).json({
        ok: true,
        skipped: true,
        reason: "Excel file unchanged (SHA-256 match) - avoided redundant DB write.",
        fileSha256,
        latestFloorScan: existingTelemetry.latestFloorScan || null,
        durationMs: Date.now() - startedAt
      });
    }

    const parsed = parseOvnWorkbook(fileBuffer);

    await batchUpsertOvnSchedule(supabaseUrl, serviceKey, parsed.sqlRows);
    await updateEuroplanSyncSnapshot(supabaseUrl, serviceKey, parsed, existingTelemetry, fileSha256);

    return res.status(200).json({
      ok: true,
      fileName: "Ovn Pro Schedule.xlsb",
      sheetName: parsed.sheetName,
      totalRowsUpserted: parsed.sqlRows.length,
      diecutSnapshotRows: parsed.diecutRawRows.length,
      latestFloorScan: parsed.telemetry.latestFloorScan,
      durationMs: Date.now() - startedAt
    });
  } catch (err) {
    console.error("[sync-ovn] Error:", err);
    return res.status(500).json({
      ok: false,
      error: err.message || String(err),
      durationMs: Date.now() - startedAt
    });
  }
};

