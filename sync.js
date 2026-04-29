import { google } from "googleapis";
import "dotenv/config";

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// ─── Source definitions ────────────────────────────────────────────────────
// Each source maps a Supabase table → a named sheet tab + master row shape.
// Use service_role keys (not anon) — they bypass RLS and guarantee reads.
//
// Yogal's table is "Signup" (PascalCase) because Prisma double-quotes model
// names in Postgres: model Signup → CREATE TABLE "Signup" → REST: /Signup
// ──────────────────────────────────────────────────────────────────────────

const SOURCES = [
  {
    sheetName: "AFK Waitlist",
    url: process.env.SUPABASE_URL_AFK_WAITLIST,
    key: process.env.SUPABASE_KEY_AFK_WAITLIST,
    table: "waitlist",
    headers: [
      "ID",
      "Name",
      "Email",
      "Phone",
      "Country Code",
      "Country",
      "Created At",
    ],
    columns: [
      "id",
      "name",
      "email",
      "phone",
      "country_code",
      "country",
      "created_at",
    ],
    toRow: (r) => [
      r.id,
      r.name,
      r.email,
      r.phone,
      r.country_code,
      r.country,
      r.created_at,
    ],
    toMaster: (r) => ({
      name: r.name ?? "",
      email: r.email ?? "",
      phone: String(r.phone ?? ""),
      source: "AFK Waitlist",
      created_at: r.created_at,
    }),
  },
  {
    sheetName: "Beat Everyday",
    url: process.env.SUPABASE_URL_BEAT_EVERYDAY,
    key: process.env.SUPABASE_KEY_BEAT_EVERYDAY,
    table: "leads_beats",
    headers: [
      "ID",
      "Name",
      "Email",
      "Phone",
      "Enjoying",
      "Scene",
      "Created At",
    ],
    columns: [
      "id",
      "name",
      "email",
      "phone",
      "enjoying",
      "scene",
      "created_at",
    ],
    toRow: (r) => [
      r.id,
      r.name,
      r.email,
      r.phone,
      r.enjoying,
      r.scene,
      r.created_at,
    ],
    toMaster: (r) => ({
      name: r.name ?? "",
      email: r.email ?? "",
      phone: r.phone ?? "",
      source: "Beat Everyday",
      created_at: r.created_at,
    }),
  },
  {
    sheetName: "AFK Extension",
    url: process.env.SUPABASE_URL_AFK_EXTENSION,
    key: process.env.SUPABASE_KEY_AFK_EXTENSION,
    table: "leads",
    headers: [
      "ID",
      "Name",
      "Phone",
      "Country Code",
      "Country",
      "Source",
      "Referral Code",
      "Created At",
    ],
    columns: [
      "id",
      "name",
      "phone",
      "country_code",
      "country",
      "source",
      "referral_code",
      "created_at",
    ],
    toRow: (r) => [
      r.id,
      r.name,
      r.phone,
      r.country_code,
      r.country,
      r.source,
      r.referral_code,
      r.created_at,
    ],
    toMaster: (r) => ({
      name: r.name ?? "",
      email: "",
      phone: r.phone ?? "",
      source: "AFK Extension",
      created_at: r.created_at,
    }),
  },
  {
    sheetName: "Yogal",
    url: process.env.SUPABASE_URL_YOGAL,
    key: process.env.SUPABASE_KEY_YOGAL,
    table: "Signup",
    headers: [
      "ID",
      "Name",
      "Email",
      "Phone",
      "Pose ID",
      "Clues Used",
      "Trigger",
      "Created At",
    ],
    columns: [
      "id",
      "name",
      "email",
      "phone",
      "poseId",
      "cluesUsed",
      "trigger",
      "createdAt",
    ],
    toRow: (r) => [
      r.id,
      r.name,
      r.email,
      r.phone,
      r.poseId,
      r.cluesUsed,
      r.trigger,
      r.createdAt,
    ],
    toMaster: (r) => ({
      name: r.name ?? "",
      email: r.email ?? "",
      phone: r.phone ?? "",
      source: "Yogal",
      created_at: r.createdAt,
    }),
  },
];

// ─── Supabase fetch ────────────────────────────────────────────────────────

async function fetchSupabase(url, key, table, columns) {
  const select = columns.join(",");
  const orderCol = columns.includes("created_at") ? "created_at" : "createdAt";
  const endpoint = `${url}/rest/v1/${table}?select=${select}&order=${orderCol}.asc`;

  const res = await fetch(endpoint, {
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
    },
  });

  if (!res.ok) {
    const body = await res.text();
    throw new Error(`[${table}] Supabase ${res.status}: ${body}`);
  }

  return res.json();
}

// ─── Google Sheets helpers ─────────────────────────────────────────────────

async function getSheets() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  // dotenv doesn't expand \n inside JSON strings — the private key arrives
  // as literal backslash-n sequences which OpenSSL can't decode.
  if (credentials.private_key) {
    credentials.private_key = credentials.private_key.replace(/\\n/g, "\n");
  }
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

async function ensureSheet(sheets, sheetName, existingTitles) {
  if (existingTitles.includes(sheetName)) return;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{ addSheet: { properties: { title: sheetName } } }],
    },
  });
}

async function writeSheet(sheets, sheetName, rows) {
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: sheetName,
  });
  if (rows.length === 0) return;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: "RAW",
    requestBody: { values: rows },
  });
}

// ─── Master deduplication ──────────────────────────────────────────────────
// Primary key: email (lowercase). Secondary key: phone (for email-less rows).
// When two entries share a key, keep the one with the earlier created_at.

function buildMaster(masterRows) {
  const byEmail = new Map();
  const byPhone = new Map();

  for (const row of masterRows) {
    const email = row.email?.trim().toLowerCase();
    const phone = row.phone?.trim();

    if (email) {
      const existing = byEmail.get(email);
      if (!existing || row.created_at < existing.created_at) {
        byEmail.set(email, row);
      }
    } else if (phone) {
      const existing = byPhone.get(phone);
      if (!existing || row.created_at < existing.created_at) {
        byPhone.set(phone, row);
      }
    }
  }

  return [...byEmail.values(), ...byPhone.values()].sort((a, b) =>
    (a.created_at ?? "").localeCompare(b.created_at ?? ""),
  );
}

// ─── Main ──────────────────────────────────────────────────────────────────

async function main() {
  if (!SPREADSHEET_ID) throw new Error("SPREADSHEET_ID is not set");
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON)
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not set");

  const sheets = await getSheets();

  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const existingTitles = meta.data.sheets.map((s) => s.properties.title);

  const allMasterRows = [];
  const results = [];

  for (const source of SOURCES) {
    if (!source.url || !source.key) {
      console.warn(`⚠  Skipping "${source.sheetName}" — env vars not set`);
      continue;
    }

    try {
      console.log(`→ Fetching ${source.sheetName}...`);
      const rows = await fetchSupabase(
        source.url,
        source.key,
        source.table,
        source.columns,
      );

      await ensureSheet(sheets, source.sheetName, existingTitles);
      existingTitles.push(source.sheetName);

      const sheetRows = [source.headers, ...rows.map(source.toRow)];
      await writeSheet(sheets, source.sheetName, sheetRows);

      rows.forEach((r) => allMasterRows.push(source.toMaster(r)));

      results.push(`✓ ${source.sheetName}: ${rows.length} rows`);
    } catch (err) {
      results.push(`✗ ${source.sheetName}: ${err.message}`);
      console.error(err.message);
    }
  }

  // Write master sheet
  const MASTER_SHEET = "Master";
  const masterRows = buildMaster(allMasterRows);

  await ensureSheet(sheets, MASTER_SHEET, existingTitles);
  const masterHeaders = ["Name", "Email", "Phone", "Source", "Created At"];
  const masterData = [
    masterHeaders,
    ...masterRows.map((r) => [
      r.name,
      r.email,
      r.phone,
      r.source,
      r.created_at,
    ]),
  ];
  await writeSheet(sheets, MASTER_SHEET, masterData);
  results.push(`✓ Master: ${masterRows.length} unique leads`);

  console.log("\n── Sync complete ──────────────────────────────");
  results.forEach((r) => console.log(r));
}

main().catch((err) => {
  console.error("Fatal:", err.message);
  process.exit(1);
});
