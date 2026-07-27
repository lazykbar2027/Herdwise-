import { Database } from "bun:sqlite";
import { join } from "node:path";

const DB_PATH = join(import.meta.dir, "herdwise.db");

export const db = new Database(DB_PATH);

// Enable WAL mode for better concurrent reads
db.exec("PRAGMA journal_mode = WAL");
db.exec("PRAGMA foreign_keys = ON");

export function initDB() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS pastures (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL UNIQUE
    );

    CREATE TABLE IF NOT EXISTS cattle (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      tag_number TEXT NOT NULL UNIQUE,
      breed TEXT NOT NULL DEFAULT '',
      sex TEXT NOT NULL CHECK(sex IN ('Bull', 'Cow', 'Steer', 'Heifer')),
      birth_date TEXT,
      pasture_id INTEGER,
      notes TEXT DEFAULT '',
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (pasture_id) REFERENCES pastures(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS breeding_records (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      cow_id INTEGER NOT NULL,
      bull_tag TEXT NOT NULL,
      breeding_date TEXT NOT NULL,
      notes TEXT DEFAULT '',
      FOREIGN KEY (cow_id) REFERENCES cattle(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS calves (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      breeding_record_id INTEGER NOT NULL,
      tag_number TEXT NOT NULL,
      sex TEXT NOT NULL CHECK(sex IN ('Bull', 'Heifer')),
      birth_date TEXT NOT NULL,
      notes TEXT DEFAULT '',
      FOREIGN KEY (breeding_record_id) REFERENCES breeding_records(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS health_records (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      cattle_id INTEGER NOT NULL,
      date TEXT NOT NULL,
      concern TEXT NOT NULL,
      resolved INTEGER NOT NULL DEFAULT 0,
      FOREIGN KEY (cattle_id) REFERENCES cattle(id) ON DELETE CASCADE
    );
  `);

  // Seed pastures if empty
  const count = db.query("SELECT COUNT(*) as c FROM pastures").get() as { c: number };
  if (count.c === 0) {
    const insert = db.prepare("INSERT INTO pastures (name) VALUES (?)");
    insert.run("North Pasture");
    insert.run("South Pasture");
    insert.run("East Pasture");
  }
}
