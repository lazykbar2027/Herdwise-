import { db } from "./db";

// ─── Types ────────────────────────────────────────────────────────────

export interface CattleListRow {
  id: number;
  tag_number: string;
  eid_tag: string | null;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_id: number | null;
  pasture_name: string | null;
  notes: string;
  status: string;
  deceased_date: string | null;
  last_health_concern: string | null;
  last_health_date: string | null;
  last_health_resolved: number | null;
  created_at: string;
  updated_at: string;
}

export interface CattleDetailRow {
  id: number;
  tag_number: string;
  eid_tag: string | null;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_id: number | null;
  pasture_name: string | null;
  notes: string;
  status: string;
  deceased_date: string | null;
  created_at: string;
  updated_at: string;
}

export interface PastureRow {
  id: number;
  name: string;
  cattle_count: number;
}

export interface HealthRecordRow {
  id: number;
  cattle_id: number;
  tag_number: string;
  date: string;
  concern: string;
  resolved: number;
}

export interface BreedingRecordRow {
  id: number;
  cow_id: number;
  cow_tag: string;
  bull_tag: string;
  breeding_date: string;
  notes: string;
  calf_tag: string | null;
  calf_sex: string | null;
  calf_birth: string | null;
}

// ─── Cattle Queries ───────────────────────────────────────────────────

const CATTLE_LIST_SELECT = `
  SELECT c.id, c.tag_number, c.eid_tag, c.breed, c.sex, c.birth_date,
         c.pasture_id, p.name as pasture_name, c.notes,
         c.status, c.deceased_date,
         (SELECT hr.concern FROM health_records hr
          WHERE hr.cattle_id = c.id ORDER BY hr.date DESC LIMIT 1) as last_health_concern,
         (SELECT hr.date FROM health_records hr
          WHERE hr.cattle_id = c.id ORDER BY hr.date DESC LIMIT 1) as last_health_date,
         (SELECT hr.resolved FROM health_records hr
          WHERE hr.cattle_id = c.id ORDER BY hr.date DESC LIMIT 1) as last_health_resolved,
         c.created_at, c.updated_at
  FROM cattle c
  LEFT JOIN pastures p ON c.pasture_id = p.id
`;

export function getCattleList(userId: number, search?: string, statusFilter?: string): CattleListRow[] {
  let sql = CATTLE_LIST_SELECT;
  sql += ` WHERE c.user_id = ?`;

  // Apply status filter
  if (statusFilter === "active") {
    sql += ` AND c.status = 'Active'`;
  } else if (statusFilter === "deceased") {
    sql += ` AND c.status = 'Deceased'`;
  }
  // "all" or undefined: no filter

  if (search && search.trim().length > 0) {
    const term = `%${search.trim()}%`;
    sql += ` AND (c.tag_number LIKE ? OR c.eid_tag LIKE ? OR c.breed LIKE ? OR c.notes LIKE ?)`;
    sql += ` ORDER BY c.tag_number`;
    return db.query(sql).all(userId, term, term, term, term) as CattleListRow[];
  }

  sql += ` ORDER BY c.tag_number`;
  return db.query(sql).all(userId) as CattleListRow[];
}

export function getCattleById(userId: number, id: number): CattleDetailRow | null {
  return db.query(`
    SELECT c.id, c.tag_number, c.eid_tag, c.breed, c.sex, c.birth_date,
           c.pasture_id, p.name as pasture_name, c.notes,
           c.status, c.deceased_date,
           c.created_at, c.updated_at
    FROM cattle c
    LEFT JOIN pastures p ON c.pasture_id = p.id
    WHERE c.id = ? AND c.user_id = ?
  `).get(id, userId) as CattleDetailRow | null;
}

export function createCattle(userId: number, data: {
  tag_number: string;
  eid_tag: string | null;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_id: number | null;
  notes: string;
}): void {
  db.run(
    `INSERT INTO cattle (user_id, tag_number, eid_tag, breed, sex, birth_date, pasture_id, notes)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
    [userId, data.tag_number, data.eid_tag, data.breed, data.sex, data.birth_date, data.pasture_id, data.notes]
  );
}

export function updateCattle(userId: number, id: number, data: {
  tag_number: string;
  eid_tag: string | null;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_id: number | null;
  notes: string;
  status: string;
  deceased_date: string | null;
}): void {
  db.run(
    `UPDATE cattle SET tag_number = ?, eid_tag = ?, breed = ?, sex = ?, birth_date = ?,
     pasture_id = ?, notes = ?, status = ?, deceased_date = ?, updated_at = datetime('now')
     WHERE id = ? AND user_id = ?`,
    [data.tag_number, data.eid_tag, data.breed, data.sex, data.birth_date, data.pasture_id, data.notes, data.status, data.deceased_date, id, userId]
  );
}

export function deleteCattle(userId: number, id: number): void {
  db.run(`DELETE FROM cattle WHERE id = ? AND user_id = ?`, [id, userId]);
}

export function getCattleInPasture(userId: number, pastureId: number, search?: string): CattleListRow[] {
  let sql = CATTLE_LIST_SELECT;
  sql += ` WHERE c.pasture_id = ? AND c.user_id = ?`;

  if (search && search.trim().length > 0) {
    const term = `%${search.trim()}%`;
    sql += ` AND (c.tag_number LIKE ? OR c.eid_tag LIKE ? OR c.breed LIKE ? OR c.notes LIKE ?)`;
    sql += ` ORDER BY c.tag_number`;
    return db.query(sql).all(pastureId, userId, term, term, term, term) as CattleListRow[];
  }

  sql += ` ORDER BY c.tag_number`;
  return db.query(sql).all(pastureId, userId) as CattleListRow[];
}

// ─── Pasture Queries ──────────────────────────────────────────────────

export function getPastures(userId: number): PastureRow[] {
  return db.query(`
    SELECT p.id, p.name, COUNT(c.id) as cattle_count
    FROM pastures p
    LEFT JOIN cattle c ON c.pasture_id = p.id
    WHERE p.user_id = ?
    GROUP BY p.id
    ORDER BY p.name
  `).all(userId) as PastureRow[];
}

export function getPastureById(userId: number, id: number): { id: number; name: string } | null {
  return db.query(`SELECT id, name FROM pastures WHERE id = ? AND user_id = ?`).get(id, userId) as { id: number; name: string } | null;
}

export function createPasture(userId: number, name: string): void {
  db.run(`INSERT INTO pastures (user_id, name) VALUES (?, ?)`, [userId, name]);
}

export function deletePasture(userId: number, id: number): void {
  db.run(`DELETE FROM pastures WHERE id = ? AND user_id = ?`, [id, userId]);
}

export function isPastureEmpty(userId: number, id: number): boolean {
  const row = db.query(`SELECT COUNT(*) as c FROM cattle WHERE pasture_id = ? AND user_id = ?`).get(id, userId) as { c: number };
  return row.c === 0;
}

export function getPastureByName(userId: number, name: string): { id: number; name: string } | null {
  return db.query(`SELECT id, name FROM pastures WHERE user_id = ? AND LOWER(name) = LOWER(?)`).get(userId, name) as { id: number; name: string } | null;
}

// ─── Health Record Queries ────────────────────────────────────────────

export function getHealthRecordsForCattle(userId: number, cattleId: number): HealthRecordRow[] {
  return db.query(`
    SELECT hr.id, hr.cattle_id, c.tag_number, hr.date, hr.concern, hr.resolved
    FROM health_records hr
    JOIN cattle c ON c.id = hr.cattle_id
    WHERE hr.cattle_id = ? AND hr.user_id = ? AND c.user_id = ?
    ORDER BY hr.date DESC
  `).all(cattleId, userId, userId) as HealthRecordRow[];
}

export function getHealthOverview(userId: number): HealthRecordRow[] {
  return db.query(`
    SELECT hr.id, hr.cattle_id, c.tag_number, hr.date, hr.concern, hr.resolved
    FROM health_records hr
    JOIN cattle c ON c.id = hr.cattle_id
    WHERE hr.user_id = ?
    ORDER BY hr.date DESC
    LIMIT 100
  `).all(userId) as HealthRecordRow[];
}

// ─── Breeding Record Queries ──────────────────────────────────────────

export function getBreedingRecordsForCattle(userId: number, cattleId: number): BreedingRecordRow[] {
  return db.query(`
    SELECT br.id, br.cow_id, cow.tag_number as cow_tag, br.bull_tag, br.breeding_date, br.notes,
           cal.tag_number as calf_tag, cal.sex as calf_sex, cal.birth_date as calf_birth
    FROM breeding_records br
    JOIN cattle cow ON cow.id = br.cow_id
    LEFT JOIN calves cal ON cal.breeding_record_id = br.id
    WHERE br.cow_id = ? AND br.user_id = ?
    ORDER BY br.breeding_date DESC
  `).all(cattleId, userId) as BreedingRecordRow[];
}

export function getBreedingOverview(userId: number): BreedingRecordRow[] {
  return getBreedingRecords(userId);
}

export function getBreedingRecords(userId: number): BreedingRecordRow[] {
  return db.query(`
    SELECT br.id, br.cow_id, cow.tag_number as cow_tag, br.bull_tag, br.breeding_date, br.notes,
           cal.tag_number as calf_tag, cal.sex as calf_sex, cal.birth_date as calf_birth
    FROM breeding_records br
    JOIN cattle cow ON cow.id = br.cow_id
    LEFT JOIN calves cal ON cal.breeding_record_id = br.id
    WHERE br.user_id = ?
    ORDER BY br.breeding_date DESC
    LIMIT 100
  `).all(userId) as BreedingRecordRow[];
}

export function getBreedingRecordById(userId: number, id: number): BreedingRecordRow | null {
  return db.query(`
    SELECT br.id, br.cow_id, cow.tag_number as cow_tag, br.bull_tag, br.breeding_date, br.notes,
           cal.tag_number as calf_tag, cal.sex as calf_sex, cal.birth_date as calf_birth
    FROM breeding_records br
    JOIN cattle cow ON cow.id = br.cow_id
    LEFT JOIN calves cal ON cal.breeding_record_id = br.id
    WHERE br.id = ? AND br.user_id = ?
  `).get(id, userId) as BreedingRecordRow | null;
}

export function createBreedingRecord(userId: number, cowId: number, bullTag: string, date: string, notes: string): void {
  db.run(
    `INSERT INTO breeding_records (user_id, cow_id, bull_tag, breeding_date, notes) VALUES (?, ?, ?, ?, ?)`,
    [userId, cowId, bullTag, date, notes]
  );
}

export function createCalf(userId: number, breedingId: number, tagNumber: string, sex: string, birthDate: string, notes: string): void {
  db.run(
    `INSERT INTO calves (user_id, breeding_record_id, tag_number, sex, birth_date, notes) VALUES (?, ?, ?, ?, ?, ?)`,
    [userId, breedingId, tagNumber, sex, birthDate, notes]
  );
}

export function deleteBreedingRecord(userId: number, id: number): void {
  db.run(`DELETE FROM breeding_records WHERE id = ? AND user_id = ?`, [id, userId]);
}

export function getHealthRecords(userId: number): HealthRecordRow[] {
  return db.query(`
    SELECT hr.id, hr.cattle_id, c.tag_number, hr.date, hr.concern, hr.resolved
    FROM health_records hr
    JOIN cattle c ON c.id = hr.cattle_id
    WHERE hr.user_id = ?
    ORDER BY hr.date DESC
    LIMIT 100
  `).all(userId) as HealthRecordRow[];
}

export function createHealthRecord(userId: number, cattleId: number, date: string, concern: string, resolved: number): void {
  db.run(
    `INSERT INTO health_records (user_id, cattle_id, date, concern, resolved) VALUES (?, ?, ?, ?, ?)`,
    [userId, cattleId, date, concern, resolved]
  );
}

export function toggleHealthRecord(userId: number, id: number): void {
  db.run(`UPDATE health_records SET resolved = CASE WHEN resolved = 0 THEN 1 ELSE 0 END WHERE id = ? AND user_id = ?`, [id, userId]);
}

export function deleteHealthRecord(userId: number, id: number): void {
  db.run(`DELETE FROM health_records WHERE id = ? AND user_id = ?`, [id, userId]);
}

// ─── Selection Helpers ──────────────────────────────────────────────

export interface CattleOption {
  id: number;
  tag_number: string;
  sex: string;
}

export function getFemaleCattle(userId: number): CattleOption[] {
  return db.query(`
    SELECT id, tag_number, sex FROM cattle
    WHERE user_id = ? AND sex IN ('Cow', 'Heifer') AND status = 'Active'
    ORDER BY tag_number
  `).all(userId) as CattleOption[];
}

export function getAllCattleOptions(userId: number): CattleOption[] {
  return db.query(`
    SELECT id, tag_number, sex, status FROM cattle
    WHERE user_id = ?
    ORDER BY tag_number
  `).all(userId) as (CattleOption & { status: string })[];
}

// ─── Import / Bulk Insert ─────────────────────────────────────────────

export function getCattleByTag(userId: number, tag: string): { id: number } | null {
  return db.query(`SELECT id FROM cattle WHERE user_id = ? AND tag_number = ?`).get(userId, tag) as { id: number } | null;
}

export function getCattleByEid(userId: number, eidTag: string): CattleListRow | null {
  const sql = CATTLE_LIST_SELECT + ` WHERE c.user_id = ? AND c.eid_tag = ?`;
  return db.query(sql).get(userId, eidTag) as CattleListRow | null;
}

export function importCattle(userId: number, records: Array<{
  tag_number: string;
  eid_tag: string | null;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_id: number | null;
  notes: string;
}>): { imported: number; skipped: Array<{ tag: string; reason: string }> } {
  const insertStmt = db.prepare(
    `INSERT INTO cattle (user_id, tag_number, eid_tag, breed, sex, birth_date, pasture_id, notes)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
  );

  let imported = 0;
  const skipped: Array<{ tag: string; reason: string }> = [];

  for (const record of records) {
    try {
      insertStmt.run(
        userId,
        record.tag_number,
        record.eid_tag,
        record.breed,
        record.sex,
        record.birth_date,
        record.pasture_id,
        record.notes
      );
      imported++;
    } catch (err: any) {
      if (err.message?.includes("UNIQUE")) {
        skipped.push({ tag: record.tag_number, reason: `Duplicate tag "${record.tag_number}" already exists` });
      } else {
        skipped.push({ tag: record.tag_number, reason: err.message || "Unknown error" });
      }
    }
  }

  return { imported, skipped };
}

// ─── Full Export Data ──────────────────────────────────────────────────

export interface ExportCattle {
  tag_number: string;
  eid_tag: string | null;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_name: string | null;
  health_status: string;
  notes: string;
  status: string;
  deceased_date: string | null;
}

export interface ExportBreeding {
  cow_tag: string;
  bull_tag: string;
  breeding_date: string;
  calf_tag: string | null;
  calf_sex: string | null;
  calf_birth_date: string | null;
  notes: string;
}

export interface ExportHealth {
  tag_number: string;
  date: string;
  concern: string;
  status: string;
}

export interface FullExportData {
  cattle: ExportCattle[];
  breeding: ExportBreeding[];
  health: ExportHealth[];
}

export function getFullExportData(userId: number): FullExportData {
  const cattle = db.query(`
    SELECT c.tag_number, c.eid_tag, c.breed, c.sex, c.birth_date,
           p.name as pasture_name,
           COALESCE(
             (SELECT hr.concern FROM health_records hr
              WHERE hr.cattle_id = c.id AND hr.resolved = 0
              ORDER BY hr.date DESC LIMIT 1),
             'Healthy'
           ) as health_status,
           c.notes, c.status, c.deceased_date
    FROM cattle c
    LEFT JOIN pastures p ON c.pasture_id = p.id
    WHERE c.user_id = ?
    ORDER BY c.tag_number
  `).all(userId) as ExportCattle[];

  const breeding = db.query(`
    SELECT cow.tag_number as cow_tag, br.bull_tag, br.breeding_date, br.notes,
           cal.tag_number as calf_tag, cal.sex as calf_sex,
           cal.birth_date as calf_birth_date
    FROM breeding_records br
    JOIN cattle cow ON cow.id = br.cow_id
    LEFT JOIN calves cal ON cal.breeding_record_id = br.id
    WHERE br.user_id = ?
    ORDER BY br.breeding_date DESC
  `).all(userId) as ExportBreeding[];

  const health = db.query(`
    SELECT c.tag_number, hr.date, hr.concern,
           CASE WHEN hr.resolved = 0 THEN 'Active' ELSE 'Resolved' END as status
    FROM health_records hr
    JOIN cattle c ON c.id = hr.cattle_id
    WHERE hr.user_id = ?
    ORDER BY hr.date DESC
  `).all(userId) as ExportHealth[];

  return { cattle, breeding, health };
}

// ─── Deceased Count ────────────────────────────────────────────────────

export function getDeceasedCattleCount(userId: number): number {
  const row = db.query(`SELECT COUNT(*) as c FROM cattle WHERE user_id = ? AND status = 'Deceased'`).get(userId) as { c: number };
  return row.c;
}
