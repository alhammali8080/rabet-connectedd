import "dotenv/config";
import express from "express";
import cors from "cors";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import Database from "better-sqlite3";
import { createClient } from "@supabase/supabase-js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const port = Number(process.env.PORT || 4000);
const app = express();
const hasSupabase = Boolean(process.env.SUPABASE_URL && process.env.SUPABASE_SERVICE_ROLE_KEY);
const supabase = hasSupabase
  ? createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY, { auth: { persistSession: false } })
  : null;

app.use(cors());
app.use(express.json({ limit: "25mb" }));

let sqlite;
if (!hasSupabase) {
  const sqlitePath = path.resolve(root, process.env.SQLITE_PATH || "./data/rabet.sqlite");
  fs.mkdirSync(path.dirname(sqlitePath), { recursive: true });
  sqlite = new Database(sqlitePath);
  sqlite.pragma("journal_mode = WAL");
  sqlite.exec(`
    create table if not exists app_state (
      id integer primary key check (id = 1),
      state text not null,
      updated_at text not null
    )
  `);
}

async function readState() {
  if (hasSupabase) {
    const { data, error } = await supabase.from("app_state").select("state").eq("id", 1).maybeSingle();
    if (error) throw error;
    return data?.state || null;
  }
  const row = sqlite.prepare("select state from app_state where id = 1").get();
  return row ? JSON.parse(row.state) : null;
}

async function writeState(state) {
  const now = new Date().toISOString();
  if (hasSupabase) {
    const { error } = await supabase.from("app_state").upsert({ id: 1, state, updated_at: now });
    if (error) throw error;
    return;
  }
  sqlite.prepare(`
    insert into app_state (id, state, updated_at) values (1, ?, ?)
    on conflict(id) do update set state = excluded.state, updated_at = excluded.updated_at
  `).run(JSON.stringify(state), now);
}

app.get("/api/health", async (_req, res) => {
  res.json({ ok: true, database: hasSupabase ? "supabase" : "sqlite", time: new Date().toISOString() });
});

app.get("/api/state", async (_req, res) => {
  try {
    const state = await readState();
    if (!state) return res.status(404).json({ state: null });
    res.json({ state });
  } catch (error) {
    console.error("GET /api/state", error);
    res.status(500).json({ error: "تعذر قراءة البيانات" });
  }
});

app.put("/api/state", async (req, res) => {
  try {
    if (!req.body || typeof req.body.state !== "object" || Array.isArray(req.body.state)) {
      return res.status(400).json({ error: "state يجب أن يكون كائناً" });
    }
    await writeState(req.body.state);
    res.json({ ok: true, updatedAt: new Date().toISOString() });
  } catch (error) {
    console.error("PUT /api/state", error);
    res.status(500).json({ error: "تعذر حفظ البيانات" });
  }
});

const dist = path.join(root, "dist");
if (fs.existsSync(dist)) {
  app.use(express.static(dist));
  app.use((req, res, next) => {
    if (req.method === "GET" && !req.path.startsWith("/api/") && req.accepts("html")) {
      return res.sendFile(path.join(dist, "index.html"));
    }
    next();
  });
}

app.listen(port, "0.0.0.0", () => {
  console.log(`رابط API يعمل على http://localhost:${port} (${hasSupabase ? "Supabase" : "SQLite محلي"})`);
});
