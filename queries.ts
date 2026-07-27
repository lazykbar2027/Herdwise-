import { db } from "./db";

// ─── Types ────────────────────────────────────────────────────────────

export interface CattleListRow {
  id: number;
  tag_number: string;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_id: number | null;
  pasture_name: string | null;
  notes: string;
  last_health_concern: string | null;
  last_health_date: string | null;
  last_health_resolved: number | null;
  created_at: string;
  updated_at: string;
}

export interface CattleDetailRow {
  id: number;
  tag_number: string;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_id: number | null;
  pasture_name: string | null;
  notes: string;
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
  SELECT c.id, c.tag_number, c.breed, c.sex, c.birth_date,
         c.pasture_id, p.name as pasture_name, c.notes,
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

export function getCattleList(search?: string): CattleListRow[] {
  let sql = CATTLE_LIST_SELECT;

  if (search && search.trim().length > 0) {
    const term = `%${search.trim()}%`;
    sql += ` WHERE c.tag_number LIKE ?1 OR c.breed LIKE ?1 OR c.notes LIKE ?1`;
    sql += ` ORDER BY c.tag_number`;
    return db.query(sql).all(term) as CattleListRow[];
  }

  sql += ` ORDER BY c.tag_number`;
  return db.query(sql).all() as CattleListRow[];
}

export function getCattleById(id: number): CattleDetailRow | null {
  return db.query(`
    SELECT c.id, c.tag_number, c.breed, c.sex, c.birth_date,
           c.pasture_id, p.name as pasture_name, c.notes,
           c.created_at, c.updated_at
    FROM cattle c
    LEFT JOIN pastures p ON c.pasture_id = p.id
    WHERE c.id = ?
  `).get(id) as CattleDetailRow | null;
}

export function createCattle(data: {
  tag_number: string;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_id: number | null;
  notes: string;
}): void {
  db.run(
    `INSERT INTO cattle (tag_number, breed, sex, birth_date, pasture_id, notes)
     VALUES (?, ?, ?, ?, ?, ?)`,
    [data.tag_number, data.breed, data.sex, data.birth_date, data.pasture_id, data.notes]
  );
}

export function updateCattle(id: number, data: {
  tag_number: string;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_id: number | null;
  notes: string;
}): void {
  db.run(
    `UPDATE cattle SET tag_number = ?, breed = ?, sex = ?, birth_date = ?,
     pasture_id = ?, notes = ?, updated_at = datetime('now')
     WHERE id = ?`,
    [data.tag_number, data.breed, data.sex, data.birth_date, data.pasture_id, data.notes, id]
  );
}

export function deleteCattle(id: number): void {
  db.run(`DELETE FROM cattle WHERE id = ?`, [id]);
}

export function getCattleInPasture(pastureId: number, search?: string): CattleListRow[] {
  let sql = CATTLE_LIST_SELECT;
  sql += ` WHERE c.pasture_id = ?`;

  if (search && search.trim().length > 0) {
    const term = `%${search.trim()}%`;
    sql += ` AND (c.tag_number LIKE ? OR c.breed LIKE ? OR c.notes LIKE ?)`;
    sql += ` ORDER BY c.tag_number`;
    return db.query(sql).all(pastureId, term, term, term) as CattleListRow[];
  }

  sql += ` ORDER BY c.tag_number`;
  return db.query(sql).all(pastureId) as CattleListRow[];
}

// ─── Pasture Queries ──────────────────────────────────────────────────

export function getPastures(): PastureRow[] {
  return db.query(`
    SELECT p.id, p.name, COUNT(c.id) as cattle_count
    FROM pastures p
    LEFT JOIN cattle c ON c.pasture_id = p.id
    GROUP BY p.id
    ORDER BY p.name
  `).all() as PastureRow[];
}

export function getPastureById(id: number): { id: number; name: string } | null {
  return db.query(`SELECT id, name FROM pastures WHERE id = ?`).get(id) as { id: number; name: string } | null;
}

export function createPasture(name: string): void {
  db.run(`INSERT INTO pastures (name) VALUES (?)`, [name]);
}

export function deletePasture(id: number): void {
  db.run(`DELETE FROM pastures WHERE id = ?`, [id]);
}

export function isPastureEmpty(id: number): boolean {
  const row = db.query(`SELECT COUNT(*) as c FROM cattle WHERE pasture_id = ?`).get(id) as { c: number };
  return row.c === 0;
}

// ─── Health Record Queries ────────────────────────────────────────────

export function getHealthRecordsForCattle(cattleId: number): HealthRecordRow[] {
  return db.query(`
    SELECT hr.id, hr.cattle_id, c.tag_number, hr.date, hr.concern, hr.resolved
    FROM health_records hr
    JOIN cattle c ON c.id = hr.cattle_id
    WHERE hr.cattle_id = ?
    ORDER BY hr.date DESC
  `).all(cattleId) as HealthRecordRow[];
}

export function getHealthOverview(): HealthRecordRow[] {
  return db.query(`
    SELECT hr.id, hr.cattle_id, c.tag_number, hr.date, hr.concern, hr.resolved
    FROM health_records hr
    JOIN cattle c ON c.id = hr.cattle_id
    ORDER BY hr.date DESC
    LIMIT 100
  `).all() as HealthRecordRow[];
}

// ─── Breeding Record Queries ──────────────────────────────────────────

export function getBreedingRecordsForCattle(cattleId: number): BreedingRecordRow[] {
  return db.query(`
    SELECT br.id, br.cow_id, cow.tag_number as cow_tag, br.bull_tag, br.breeding_date, br.notes,
           cal.tag_number as calf_tag, cal.sex as calf_sex, cal.birth_date as calf_birth
    FROM breeding_records br
    JOIN cattle cow ON cow.id = br.cow_id
    LEFT JOIN calves cal ON cal.breeding_record_id = br.id
    WHERE br.cow_id = ?
    ORDER BY br.breeding_date DESC
  `).all(cattleId) as BreedingRecordRow[];
}

export function getBreedingOverview(): BreedingRecordRow[] {
  return getBreedingRecords();
}

export function getBreedingRecords(): BreedingRecordRow[] {
  return db.query(`
    SELECT br.id, br.cow_id, cow.tag_number as cow_tag, br.bull_tag, br.breeding_date, br.notes,
           cal.tag_number as calf_tag, cal.sex as calf_sex, cal.birth_date as calf_birth
    FROM breeding_records br
    JOIN cattle cow ON cow.id = br.cow_id
    LEFT JOIN calves cal ON cal.breeding_record_id = br.id
    ORDER BY br.breeding_date DESC
    LIMIT 100
  `).all() as BreedingRecordRow[];
}

export function getBreedingRecordById(id: number): BreedingRecordRow | null {
  return db.query(`
    SELECT br.id, br.cow_id, cow.tag_number as cow_tag, br.bull_tag, br.breeding_date, br.notes,
           cal.tag_number as calf_tag, cal.sex as calf_sex, cal.birth_date as calf_birth
    FROM breeding_records br
    JOIN cattle cow ON cow.id = br.cow_id
    LEFT JOIN calves cal ON cal.breeding_record_id = br.id
    WHERE br.id = ?
  `).get(id) as BreedingRecordRow | null;
}

export function createBreedingRecord(cowId: number, bullTag: string, date: string, notes: string): void {
  db.run(
    `INSERT INTO breeding_records (cow_id, bull_tag, breeding_date, notes) VALUES (?, ?, ?, ?)`,
    [cowId, bullTag, date, notes]
  );
}

export function createCalf(breedingId: number, tagNumber: string, sex: string, birthDate: string, notes: string): void {
  db.run(
    `INSERT INTO calves (breeding_record_id, tag_number, sex, birth_date, notes) VALUES (?, ?, ?, ?, ?)`,
    [breedingId, tagNumber, sex, birthDate, notes]
  );
}

export function deleteBreedingRecord(id: number): void {
  db.run(`DELETE FROM breeding_records WHERE id = ?`, [id]);
}

export function getHealthRecords(): HealthRecordRow[] {
  return db.query(`
    SELECT hr.id, hr.cattle_id, c.tag_number, hr.date, hr.concern, hr.resolved
    FROM health_records hr
    JOIN cattle c ON c.id = hr.cattle_id
    ORDER BY hr.date DESC
    LIMIT 100
  `).all() as HealthRecordRow[];
}

export function createHealthRecord(cattleId: number, date: string, concern: string, resolved: number): void {
  db.run(
    `INSERT INTO health_records (cattle_id, date, concern, resolved) VALUES (?, ?, ?, ?)`,
    [cattleId, date, concern, resolved]
  );
}

export function toggleHealthRecord(id: number): void {
  db.run(`UPDATE health_records SET resolved = CASE WHEN resolved = 0 THEN 1 ELSE 0 END WHERE id = ?`, [id]);
}

export function deleteHealthRecord(id: number): void {
  db.run(`DELETE FROM health_records WHERE id = ?`, [id]);
}

// ─── Selection Helpers ──────────────────────────────────────────────

export interface CattleOption {
  id: number;
  tag_number: string;
  sex: string;
}

export function getFemaleCattle(): CattleOption[] {
  return db.query(`
    SELECT id, tag_number, sex FROM cattle
    WHERE sex IN ('Cow', 'Heifer')
    ORDER BY tag_number
  `).all() as CattleOption[];
}

export function getAllCattleOptions(): CattleOption[] {
  return db.query(`
    SELECT id, tag_number, sex FROM cattle
    ORDER BY tag_number
  `).all() as CattleOption[];
}

// ─── Import / Bulk Insert ─────────────────────────────────────────────

export function getCattleByTag(tag: string): { id: number } | null {
  return db.query(`SELECT id FROM cattle WHERE tag_number = ?`).get(tag) as { id: number } | null;
}

export function importCattle(records: Array<{
  tag_number: string;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_id: number | null;
  notes: string;
}>): { imported: number; skipped: Array<{ tag: string; reason: string }> } {
  const insertStmt = db.prepare(
    `INSERT INTO cattle (tag_number, breed, sex, birth_date, pasture_id, notes)
     VALUES (?, ?, ?, ?, ?, ?)`
  );

  let imported = 0;
  const skipped: Array<{ tag: string; reason: string }> = [];

  for (const record of records) {
    try {
      insertStmt.run(
        record.tag_number,
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
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture_name: string | null;
  health_status: string;
  notes: string;
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

export function getFullExportData(): FullExportData {
  const cattle = db.query(`
    SELECT c.tag_number, c.breed, c.sex, c.birth_date,
           p.name as pasture_name,
           COALESCE(
             (SELECT hr.concern FROM health_records hr
              WHERE hr.cattle_id = c.id AND hr.resolved = 0
              ORDER BY hr.date DESC LIMIT 1),
             'Healthy'
           ) as health_status,
           c.notes
    FROM cattle c
    LEFT JOIN pastures p ON c.pasture_id = p.id
    ORDER BY c.tag_number
  `).all() as ExportCattle[];

  const breeding = db.query(`
    SELECT cow.tag_number as cow_tag, br.bull_tag, br.breeding_date, br.notes,
           cal.tag_number as calf_tag, cal.sex as calf_sex,
           cal.birth_date as calf_birth_date
    FROM breeding_records br
    JOIN cattle cow ON cow.id = br.cow_id
    LEFT JOIN calves cal ON cal.breeding_record_id = br.id
    ORDER BY br.breeding_date DESC
  `).all() as ExportBreeding[];

  const health = db.query(`
    SELECT c.tag_number, hr.date, hr.concern,
           CASE WHEN hr.resolved = 0 THEN 'Active' ELSE 'Resolved' END as status
    FROM health_records hr
    JOIN cattle c ON c.id = hr.cattle_id
    ORDER BY hr.date DESC
  `).all() as ExportHealth[];

  return { cattle, breeding, health };
}
