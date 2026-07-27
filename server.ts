import { db, initDB } from "./db";
import { renderPage, escapeHTML } from "./render";
import {
  getCattleList, getCattleById, createCattle, updateCattle, deleteCattle,
  getCattleInPasture,
  getPastures, getPastureById, createPasture, deletePasture, isPastureEmpty,
  getHealthRecordsForCattle, getHealthOverview, getHealthRecords,
  createHealthRecord, toggleHealthRecord, deleteHealthRecord,
  getBreedingRecordsForCattle, getBreedingOverview, getBreedingRecords,
  getBreedingRecordById, createBreedingRecord, createCalf, deleteBreedingRecord,
  getFemaleCattle, getAllCattleOptions,
  getFullExportData,
  importCattle, getCattleByTag,
} from "./queries";
import { readFileSync, existsSync } from "node:fs";
import { join, extname } from "node:path";
import ExcelJS from "exceljs";
import type { CattleListRow, PastureRow, HealthRecordRow, BreedingRecordRow, CattleOption } from "./queries";

// Initialize the database
initDB();

const PUBLIC_DIR = join(import.meta.dir, "public");

const MIME_TYPES: Record<string, string> = {
  ".html": "text/html",
  ".css": "text/css",
  ".js": "application/javascript",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
};

// ─── Helpers ──────────────────────────────────────────────────────────

function serveStatic(pathname: string): Response | null {
  const filePath = join(PUBLIC_DIR, pathname === "/" ? "index.html" : pathname);
  if (!existsSync(filePath)) return null;
  const ext = extname(filePath);
  const mime = MIME_TYPES[ext] || "application/octet-stream";
  return new Response(readFileSync(filePath), {
    headers: { "Content-Type": mime },
  });
}

function redirect(url: string): Response {
  return new Response(null, {
    status: 302,
    headers: { Location: url },
  });
}

function htmlPage(title: string, body: string, status: number = 200): Response {
  return new Response(renderPage(title, body), {
    status,
    headers: { "Content-Type": "text/html; charset=utf-8" },
  });
}

function notFound(): Response {
  return htmlPage("Not Found", `<div class="empty-state"><p>Page not found.</p><a href="/cattle">← Back to Cattle</a></div>`, 404);
}

function sexBadge(sex: string): string {
  const colorMap: Record<string, string> = {
    Bull: "#8B5E3C",
    Cow: "#3B7DD8",
    Heifer: "#5D9B4C",
    Steer: "#888888",
  };
  const color = colorMap[sex] || "#888";
  return `<span class="sex-badge" style="background:${color}">${escapeHTML(sex)}</span>`;
}

function healthBadge(concern: string | null, date: string | null, resolved: number | null): string {
  if (!concern) return `<span class="health-ok">No concerns</span>`;
  if (resolved) return `<span class="health-resolved">${escapeHTML(concern)} (resolved)</span>`;
  return `<span class="health-active">${escapeHTML(concern)} — ${escapeHTML(date || "")}</span>`;
}

function formatDate(d: string | null): string {
  if (!d) return "—";
  return d;
}

function cattleTableRows(cattle: CattleListRow[], showEdit: boolean = true): string {
  if (cattle.length === 0) return "";
  return cattle.map(c => `
    <tr>
      <td><a href="/cattle/${c.id}" class="tag-link">${escapeHTML(c.tag_number)}</a></td>
      <td>${escapeHTML(c.breed || "—")}</td>
      <td>${sexBadge(c.sex)}</td>
      <td>${c.pasture_name ? `<a href="/pastures/${c.pasture_id}">${escapeHTML(c.pasture_name)}</a>` : "—"}</td>
      <td>${healthBadge(c.last_health_concern, c.last_health_date, c.last_health_resolved)}</td>
      ${showEdit ? `<td><a href="/cattle/${c.id}/edit" class="btn btn-sm btn-outline">Edit</a></td>` : ""}
    </tr>
  `).join("");
}

function pastureOptions(selectedId: number | null = null): string {
  const pastures = db.query("SELECT id, name FROM pastures ORDER BY name").all() as { id: number; name: string }[];
  return pastures.map(p =>
    `<option value="${p.id}"${selectedId === p.id ? " selected" : ""}>${escapeHTML(p.name)}</option>`
  ).join("");
}

function femaleCattleOptions(selectedId: number | null = null): string {
  const cattle = getFemaleCattle();
  return cattle.map(c =>
    `<option value="${c.id}"${selectedId === c.id ? " selected" : ""}>${escapeHTML(c.tag_number)} (${escapeHTML(c.sex)})</option>`
  ).join("");
}

function allCattleOptions(selectedId: number | null = null): string {
  const cattle = getAllCattleOptions();
  return cattle.map(c =>
    `<option value="${c.id}"${selectedId === c.id ? " selected" : ""}>${escapeHTML(c.tag_number)} (${escapeHTML(c.sex)})</option>`
  ).join("");
}

function cattleFormHTML(
  action: string,
  defaults: { tag_number?: string; breed?: string; sex?: string; birth_date?: string; pasture_id?: number | null; notes?: string; } = {},
  error?: string,
  includeDelete: boolean = false,
  deleteAction?: string
): string {
  const isEdit = !!includeDelete;

  return `
    ${error ? `<div class="alert alert-error">${escapeHTML(error)}</div>` : ""}
    <form method="POST" action="${action}" class="cattle-form">
      <label for="tag_number">Tag Number *</label>
      <input type="text" id="tag_number" name="tag_number" required maxlength="50"
             value="${escapeHTML(defaults.tag_number || "")}" autofocus>

      <label for="breed">Breed</label>
      <input type="text" id="breed" name="breed" maxlength="100" placeholder="e.g. Angus, Hereford"
             value="${escapeHTML(defaults.breed || "")}">

      <label for="sex">Sex *</label>
      <select id="sex" name="sex" required>
        <option value="">— Select —</option>
        <option value="Bull"${defaults.sex === "Bull" ? " selected" : ""}>Bull</option>
        <option value="Cow"${defaults.sex === "Cow" ? " selected" : ""}>Cow</option>
        <option value="Steer"${defaults.sex === "Steer" ? " selected" : ""}>Steer</option>
        <option value="Heifer"${defaults.sex === "Heifer" ? " selected" : ""}>Heifer</option>
      </select>

      <label for="birth_date">Birth Date</label>
      <input type="date" id="birth_date" name="birth_date"
             value="${escapeHTML(defaults.birth_date || "")}">

      <label for="pasture_id">Pasture</label>
      <select id="pasture_id" name="pasture_id">
        <option value="">— None —</option>
        ${pastureOptions(defaults.pasture_id ?? null)}
      </select>

      <label for="notes">Notes</label>
      <textarea id="notes" name="notes" maxlength="500" placeholder="General notes about this animal">${escapeHTML(defaults.notes || "")}</textarea>

      <div class="form-actions">
        <button type="submit" class="btn">${isEdit ? "Update Cattle" : "Save Cattle"}</button>
        <a href="/cattle" class="btn btn-cancel">Cancel</a>
      </div>
    </form>

    ${includeDelete && deleteAction ? `
    <form method="POST" action="${deleteAction}" class="delete-form" onsubmit="return confirm('Delete this cattle? This cannot be undone.')">
      <button type="submit" class="btn btn-danger">Delete This Cattle</button>
    </form>` : ""}

    <p class="mt-1"><a href="/cattle">← Back to Cattle</a></p>
  `;
}

// ─── Import Helpers ────────────────────────────────────────────────────

// In-memory storage for previewed import data (single-user app)
let previewData: ImportRecord[] = [];

interface ImportRecord {
  tag_number: string;
  breed: string;
  sex: string;
  birth_date: string | null;
  pasture: string;
  notes: string;
  errors: string[];
}

// Flexible column name mapping
function normalizeHeader(h: string): string {
  const lower = h.toLowerCase().trim();
  // Remove non-alphanumeric chars for comparison
  const clean = lower.replace(/[^a-z0-9]/g, "");
  if (/tag|tagnumber|id|cattle.?id/.test(clean) && !/pasture/.test(clean)) return "tag_number";
  if (/breed/.test(clean)) return "breed";
  if (/sex|gender/.test(clean)) return "sex";
  if (/birth|dob|birthdate|date.?of.?birth/.test(clean)) return "birth_date";
  if (/pasture|pasturename|location/.test(clean)) return "pasture";
  if (/note|comment|remark/.test(clean)) return "notes";
  return "";
}

function parseCSV(text: string): string[][] {
  const rows: string[][] = [];
  const lines = text.split(/\r?\n/).filter(line => line.trim() !== "");
  for (const line of lines) {
    const row: string[] = [];
    let current = "";
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i]!;
      if (inQuotes) {
        if (ch === '"') {
          if (i + 1 < line.length && line[i + 1] === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          current += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === ",") {
          row.push(current.trim());
          current = "";
        } else {
          current += ch;
        }
      }
    }
    row.push(current.trim());
    rows.push(row);
  }
  return rows;
}

async function parseFile(buffer: ArrayBuffer, filename: string): Promise<string[][]> {
  const lower = filename.toLowerCase();
  if (lower.endsWith(".csv")) {
    const text = new TextDecoder().decode(buffer);
    return parseCSV(text);
  }
  if (lower.endsWith(".xlsx")) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) return [];
    const rows: string[][] = [];
    worksheet.eachRow((row, rowNumber) => {
      const values: string[] = [];
      row.eachCell((cell) => {
        let val = "";
        if (cell.value !== null && cell.value !== undefined) {
          if (cell.value instanceof Date) {
            val = cell.value.toISOString().split("T")[0]!;
          } else {
            val = String(cell.value);
          }
        }
        values.push(val);
      });
      rows.push(values);
    });
    return rows;
  }
  return [];
}

function validateRecord(record: ImportRecord): string[] {
  const errors: string[] = [];
  if (!record.tag_number || record.tag_number.trim() === "") {
    errors.push("Tag number is required");
  }
  const validSexes = ["Bull", "Cow", "Steer", "Heifer"];
  if (record.sex && !validSexes.includes(record.sex)) {
    errors.push(`Invalid sex "${record.sex}" — must be Bull, Cow, Steer, or Heifer`);
  }
  return errors;
}

function normalizeSex(raw: string): string {
  const cleaned = raw.trim();
  const lower = cleaned.toLowerCase();
  // Only return valid values; invalid ones stay as-is for validation to catch
  const valid = ["bull", "cow", "steer", "heifer"];
  if (valid.includes(lower)) {
    return cleaned.charAt(0).toUpperCase() + cleaned.slice(1).toLowerCase();
  }
  return cleaned; // pass through for validation
}

// ─── Cattle List ──────────────────────────────────────────────────────

function handleCattleList(search?: string, addedTag?: string | null, importSummary?: { imported: number; skipped: number; skippedDetails?: string }): Response {
  const cattle = getCattleList(search);

  let banner = "";
  if (addedTag) {
    banner = `<div class="alert alert-success">✅ Cattle <strong>${escapeHTML(addedTag)}</strong> added successfully!</div>`;
  }

  if (importSummary) {
    banner += `<div class="alert alert-success">
      📥 Import complete: <strong>${importSummary.imported}</strong> cattle imported${importSummary.skipped > 0 ? `, <strong>${importSummary.skipped}</strong> skipped` : ""}.
    </div>`;
    if (importSummary.skippedDetails) {
      const details = importSummary.skippedDetails.split("||").map(d => escapeHTML(d));
      banner += `<div class="alert alert-error" style="margin-top: 0.5rem;">
        <strong>Skipped rows:</strong><br>${details.join("<br>")}
      </div>`;
    }
  }

  let body = "";

  // Search bar and add button
  body += `
    <div class="list-header">
      <div class="list-header-top">
        <a href="/cattle/add" class="btn">+ Add Cattle</a>
        <span class="cattle-count">Showing <strong>${cattle.length}</strong> cattle</span>
      </div>
      <form method="GET" action="/cattle" class="search-form">
        <input type="search" name="search" placeholder="Search by tag, breed, or notes…"
               value="${escapeHTML(search || "")}" class="search-input">
        <button type="submit" class="btn btn-sm">Search</button>
        ${search ? `<a href="/cattle" class="btn btn-sm btn-outline">Clear</a>` : ""}
      </form>
    </div>
  `;

  body += banner;

  if (cattle.length === 0) {
    body += `
      <div class="empty-state">
        <p>${search ? "No cattle match your search." : "No cattle recorded yet."}</p>
        ${!search ? `<a href="/cattle/add" class="btn">Add Your First Cattle</a>` : ""}
      </div>
    `;
  } else {
    body += `
      <div class="table-wrapper">
        <table>
          <thead>
            <tr>
              <th>Tag #</th>
              <th>Breed</th>
              <th>Sex</th>
              <th>Pasture</th>
              <th>Health</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            ${cattleTableRows(cattle)}
          </tbody>
        </table>
      </div>
    `;
  }

  return htmlPage("Cattle", body);
}

// ─── Cattle Add Form ──────────────────────────────────────────────────

function handleCattleAddForm(error?: string, defaults?: Record<string, string>): Response {
  return htmlPage("Add Cattle", cattleFormHTML("/cattle", {
    tag_number: defaults?.tag_number || "",
    breed: defaults?.breed || "",
    sex: defaults?.sex || "",
    birth_date: defaults?.birth_date || "",
    pasture_id: defaults?.pasture_id ? parseInt(defaults.pasture_id, 10) : null,
    notes: defaults?.notes || "",
  }, error));
}

// ─── Cattle Create ────────────────────────────────────────────────────

async function handleCattleCreate(req: Request): Promise<Response> {
  const formData = await req.formData();
  const tag_number = formData.get("tag_number")?.toString().trim();
  const breed = formData.get("breed")?.toString().trim() || "";
  const sex = formData.get("sex")?.toString().trim();
  const birth_date = formData.get("birth_date")?.toString().trim() || null;
  const pasture_id_raw = formData.get("pasture_id")?.toString().trim();
  const notes = formData.get("notes")?.toString().trim() || "";

  // Validate required fields
  if (!tag_number || !sex) {
    return handleCattleAddForm("Tag number and sex are required.", {
      tag_number: tag_number || "",
      breed,
      sex: sex || "",
      birth_date: birth_date || "",
      pasture_id: pasture_id_raw || "",
      notes,
    });
  }

  const validSexes = ["Bull", "Cow", "Steer", "Heifer"];
  if (!validSexes.includes(sex)) {
    return handleCattleAddForm("Invalid sex value.", {
      tag_number,
      breed,
      sex,
      birth_date: birth_date || "",
      pasture_id: pasture_id_raw || "",
      notes,
    });
  }

  const pasture_id = pasture_id_raw ? parseInt(pasture_id_raw, 10) : null;

  try {
    createCattle({ tag_number, breed, sex, birth_date, pasture_id, notes });
  } catch (err: any) {
    if (err.message?.includes("UNIQUE")) {
      return handleCattleAddForm(`A cattle with tag number "${tag_number}" already exists.`, {
        tag_number: "",
        breed,
        sex,
        birth_date: birth_date || "",
        pasture_id: pasture_id_raw || "",
        notes,
      });
    }
    throw err;
  }

  return redirect(`/cattle?added=${encodeURIComponent(tag_number)}`);
}

// ─── Cattle Detail ────────────────────────────────────────────────────

function handleCattleDetail(id: number): Response {
  const cattle = getCattleById(id);
  if (!cattle) return notFound();

  const healthRecords = getHealthRecordsForCattle(id);
  const breedingRecords = getBreedingRecordsForCattle(id);

  let body = "";

  // Info card
  body += `<div class="detail-card">`;
  body += `<div class="detail-header">`;
  body += `<h2>${escapeHTML(cattle.tag_number)} ${sexBadge(cattle.sex)}</h2>`;
  body += `<div class="detail-actions">`;
  body += `<a href="/cattle/${id}/edit" class="btn btn-sm">Edit</a>`;
  body += `<a href="/cattle" class="btn btn-sm btn-outline">← Back</a>`;
  body += `</div></div>`;

  body += `<div class="detail-grid">`;
  body += `<div class="detail-item"><span class="detail-label">Breed</span><span class="detail-value">${escapeHTML(cattle.breed || "—")}</span></div>`;
  body += `<div class="detail-item"><span class="detail-label">Birth Date</span><span class="detail-value">${formatDate(cattle.birth_date)}</span></div>`;
  body += `<div class="detail-item"><span class="detail-label">Pasture</span><span class="detail-value">${cattle.pasture_name ? `<a href="/pastures/${cattle.pasture_id}">${escapeHTML(cattle.pasture_name)}</a>` : "—"}</span></div>`;
  body += `</div>`;

  if (cattle.notes) {
    body += `<div class="detail-notes"><span class="detail-label">Notes</span><p>${escapeHTML(cattle.notes)}</p></div>`;
  }

  // Quick action buttons
  body += `<div class="detail-actions mt-1">`;
  if (cattle.sex === "Cow" || cattle.sex === "Heifer") {
    body += `<a href="/breeding?form=1" class="btn btn-sm">🐄 Record Breeding</a>`;
  }
  body += `<a href="/health?form=1" class="btn btn-sm">🩺 Add Health Concern</a>`;
  body += `</div>`;

  body += `</div>`; // close detail-card

  // Health records section
  body += `<div class="detail-card">`;
  body += `<h3>🩺 Health Records</h3>`;
  if (healthRecords.length === 0) {
    body += `<p class="text-muted">No health records yet. <a href="/health?form=1">Add one →</a></p>`;
  } else {
    body += `<div class="table-wrapper"><table>`;
    body += `<thead><tr><th>Date</th><th>Concern</th><th>Status</th></tr></thead><tbody>`;
    body += healthRecords.map(r => `
      <tr>
        <td>${escapeHTML(r.date)}</td>
        <td>${escapeHTML(r.concern)}</td>
        <td>${r.resolved ? `<span class="health-resolved">✅ Resolved</span>` : `<span class="health-active">⚠️ Active</span>`}</td>
      </tr>
    `).join("");
    body += `</tbody></table></div>`;
  }
  body += `</div>`;

  // Breeding records section
  body += `<div class="detail-card">`;
  body += `<h3>🐄 Breeding Records</h3>`;
  if (cattle.sex === "Bull" || cattle.sex === "Steer") {
    body += `<p class="text-muted">Breeding records are only tracked for female cattle (Cows and Heifers).</p>`;
  } else if (breedingRecords.length === 0) {
    body += `<p class="text-muted">No breeding records yet. <a href="/breeding?form=1">Record one →</a></p>`;
  } else {
    body += `<div class="table-wrapper"><table>`;
    body += `<thead><tr><th>Bull Tag</th><th>Breeding Date</th><th>Calf Tag</th><th>Calf Sex</th><th>Calf Birth</th></tr></thead><tbody>`;
    body += breedingRecords.map(r => `
      <tr>
        <td>${escapeHTML(r.bull_tag)}</td>
        <td>${escapeHTML(r.breeding_date)}</td>
        <td>${r.calf_tag ? escapeHTML(r.calf_tag) : "—"}</td>
        <td>${r.calf_sex ? (r.calf_sex === "Heifer" ? "🐮 " : "🐂 ") + escapeHTML(r.calf_sex) : "—"}</td>
        <td>${r.calf_birth ? escapeHTML(r.calf_birth) : "—"}</td>
      </tr>
    `).join("");
    body += `</tbody></table></div>`;
  }
  body += `</div>`;

  return htmlPage(cattle.tag_number, body);
}

// ─── Cattle Edit Form ─────────────────────────────────────────────────

function handleCattleEditForm(id: number, error?: string): Response {
  const cattle = getCattleById(id);
  if (!cattle) return notFound();

  return htmlPage(`Edit ${cattle.tag_number}`, cattleFormHTML(
    `/cattle/${id}`,
    {
      tag_number: cattle.tag_number,
      breed: cattle.breed,
      sex: cattle.sex,
      birth_date: cattle.birth_date || "",
      pasture_id: cattle.pasture_id,
      notes: cattle.notes,
    },
    error,
    true,
    `/cattle/${id}/delete`
  ));
}

// ─── Cattle Update ────────────────────────────────────────────────────

async function handleCattleUpdate(req: Request, id: number): Promise<Response> {
  const cattle = getCattleById(id);
  if (!cattle) return notFound();

  const formData = await req.formData();
  const tag_number = formData.get("tag_number")?.toString().trim();
  const breed = formData.get("breed")?.toString().trim() || "";
  const sex = formData.get("sex")?.toString().trim();
  const birth_date = formData.get("birth_date")?.toString().trim() || null;
  const pasture_id_raw = formData.get("pasture_id")?.toString().trim();
  const notes = formData.get("notes")?.toString().trim() || "";

  if (!tag_number || !sex) {
    return handleCattleEditForm(id, "Tag number and sex are required.");
  }

  const validSexes = ["Bull", "Cow", "Steer", "Heifer"];
  if (!validSexes.includes(sex)) {
    return handleCattleEditForm(id, "Invalid sex value.");
  }

  const pasture_id = pasture_id_raw ? parseInt(pasture_id_raw, 10) : null;

  try {
    updateCattle(id, { tag_number, breed, sex, birth_date, pasture_id, notes });
  } catch (err: any) {
    if (err.message?.includes("UNIQUE")) {
      return handleCattleEditForm(id, `Tag number "${tag_number}" is already in use by another animal.`);
    }
    throw err;
  }

  return redirect(`/cattle?updated=${encodeURIComponent(tag_number)}`);
}

// ─── Cattle Delete ────────────────────────────────────────────────────

function handleCattleDelete(id: number): Response {
  const cattle = getCattleById(id);
  if (!cattle) {
    return redirect("/cattle");
  }
  deleteCattle(id);
  return redirect(`/cattle?deleted=${encodeURIComponent(cattle.tag_number)}`);
}

// ─── Pasture List ─────────────────────────────────────────────────────

function handlePasturesList(): Response {
  const pastures = getPastures();

  let body = "";

  body += `
    <div class="list-header">
      <div class="list-header-top">
        <span class="cattle-count">Showing <strong>${pastures.length}</strong> pastures</span>
      </div>
    </div>
  `;

  body += `
    <form method="POST" action="/pastures" class="inline-form">
      <label for="pasture_name">Add Pasture</label>
      <div class="inline-form-row">
        <input type="text" id="pasture_name" name="name" required maxlength="100" placeholder="Pasture name">
        <button type="submit" class="btn">Add</button>
      </div>
    </form>
  `;

  if (pastures.length === 0) {
    body += `<div class="empty-state"><p>No pastures set up yet.</p></div>`;
  } else {
    body += `<div class="table-wrapper"><table>`;
    body += `<thead><tr><th>Pasture</th><th>Cattle Count</th><th>Actions</th></tr></thead><tbody>`;
    body += pastures.map(p => `
      <tr>
        <td><a href="/pastures/${p.id}" class="tag-link">${escapeHTML(p.name)}</a></td>
        <td>${p.cattle_count}</td>
        <td>
          <form method="POST" action="/pastures/${p.id}/delete" style="display:inline"
                onsubmit="return confirm('Delete this pasture?')">
            <button type="submit" class="btn btn-sm btn-danger" ${p.cattle_count > 0 ? "disabled title=\"Cannot delete: cattle are assigned\"" : ""}>Delete</button>
          </form>
        </td>
      </tr>
    `).join("");
    body += `</tbody></table></div>`;
  }

  return htmlPage("Pastures", body);
}

// ─── Pasture Create ───────────────────────────────────────────────────

async function handlePastureCreate(req: Request): Promise<Response> {
  const formData = await req.formData();
  const name = formData.get("name")?.toString().trim();

  if (!name) {
    return redirect("/pastures");
  }

  try {
    createPasture(name);
  } catch (err: any) {
    // UNIQUE constraint — duplicate name, just redirect silently
    if (err.message?.includes("UNIQUE")) {
      return redirect("/pastures");
    }
    throw err;
  }

  return redirect("/pastures");
}

// ─── Pasture Delete ───────────────────────────────────────────────────

function handlePastureDelete(id: number): Response {
  if (!isPastureEmpty(id)) {
    // Show error — can't delete pasture with cattle
    const pastures = getPastures();
    const pasture = getPastureById(id);
    const name = pasture?.name || "Unknown";

    let body = `<div class="alert alert-error">Cannot delete "${escapeHTML(name)}": ${pastures.find(p => p.id === id)?.cattle_count || 0} cattle are still assigned to this pasture.</div>`;
    body += `<p class="mt-1"><a href="/pastures">← Back to Pastures</a></p>`;
    return htmlPage("Error", body);
  }

  deletePasture(id);
  return redirect("/pastures");
}

// ─── Pasture Detail ───────────────────────────────────────────────────

function handlePastureDetail(id: number, search?: string): Response {
  const pasture = getPastureById(id);
  if (!pasture) return notFound();

  const cattle = getCattleInPasture(id, search);

  let body = "";

  body += `
    <div class="list-header">
      <div class="list-header-top">
        <span class="cattle-count"><strong>${cattle.length}</strong> cattle in ${escapeHTML(pasture.name)}</span>
        <a href="/pastures" class="btn btn-sm btn-outline">← All Pastures</a>
      </div>
      <form method="GET" action="/pastures/${id}" class="search-form">
        <input type="search" name="search" placeholder="Search cattle in this pasture…"
               value="${escapeHTML(search || "")}" class="search-input">
        <button type="submit" class="btn btn-sm">Search</button>
        ${search ? `<a href="/pastures/${id}" class="btn btn-sm btn-outline">Clear</a>` : ""}
      </form>
    </div>
  `;

  if (cattle.length === 0) {
    body += `<div class="empty-state"><p>${search ? "No cattle match your search." : "No cattle in this pasture."}</p></div>`;
  } else {
    body += `
      <div class="table-wrapper">
        <table>
          <thead>
            <tr>
              <th>Tag #</th>
              <th>Breed</th>
              <th>Sex</th>
              <th>Health</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            ${cattleTableRows(cattle)}
          </tbody>
        </table>
      </div>
    `;
  }

  return htmlPage(pasture.name, body);
}

// ─── Health Overview ──────────────────────────────────────────────────

function handleHealthOverview(showForm: boolean = false, error?: string, defaults?: Record<string, string>): Response {
  const records = getHealthRecords();

  let body = "";

  // "Add Health Record" button
  body += `
    <div class="list-header">
      <div class="list-header-top">
        <a href="/health?form=1" class="btn">+ Add Health Record</a>
        <span class="cattle-count"><strong>${records.length}</strong> health record${records.length !== 1 ? "s" : ""}</span>
      </div>
    </div>
  `;

  // Show inline form if requested
  if (showForm) {
    const cattleCount = getAllCattleOptions().length;
    if (cattleCount === 0) {
      body += `<div class="alert alert-error">No cattle recorded yet. <a href="/cattle/add">Add cattle</a> first.</div>`;
    } else {
      body += `
        ${error ? `<div class="alert alert-error">${escapeHTML(error)}</div>` : ""}
        <form method="POST" action="/health">
          <h3>Add Health Record</h3>
          <label for="cattle_id">Cattle *</label>
          <select id="cattle_id" name="cattle_id" required>
            <option value="">— Select Cattle —</option>
            ${allCattleOptions(defaults?.cattle_id ? parseInt(defaults.cattle_id, 10) : null)}
          </select>

          <label for="date">Date *</label>
          <input type="date" id="date" name="date" required
                 value="${escapeHTML(defaults?.date || "")}">

          <label for="concern">Concern *</label>
          <textarea id="concern" name="concern" required maxlength="500"
                    placeholder="Describe the health concern...">${escapeHTML(defaults?.concern || "")}</textarea>

          <label class="checkbox-label">
            <input type="checkbox" name="resolved" value="1"${defaults?.resolved === "1" ? " checked" : ""}>
            Resolved
          </label>

          <div class="form-actions">
            <button type="submit" class="btn">Save Health Record</button>
            <a href="/health" class="btn btn-cancel">Cancel</a>
          </div>
        </form>
      `;
    }
  }

  // Table of records
  if (records.length === 0) {
    body += `<div class="empty-state"><p>No health records yet.</p></div>`;
  } else {
    body += `<div class="table-wrapper"><table>`;
    body += `<thead><tr>
      <th>Cattle Tag</th><th>Date</th><th>Concern</th><th>Status</th><th>Actions</th>
    </tr></thead><tbody>`;
    body += records.map(r => {
      const statusBadge = r.resolved
        ? `<span class="health-resolved">✅ Resolved</span>`
        : `<span class="health-active">⚠️ Active</span>`;

      const toggleLabel = r.resolved ? "Reopen" : "Resolve";

      return `
        <tr>
          <td><a href="/cattle/${r.cattle_id}" class="tag-link">${escapeHTML(r.tag_number)}</a></td>
          <td>${escapeHTML(r.date)}</td>
          <td>${escapeHTML(r.concern)}</td>
          <td>${statusBadge}</td>
          <td>
            <form method="POST" action="/health/${r.id}/toggle" style="display:inline">
              <button type="submit" class="btn btn-sm btn-outline">${toggleLabel}</button>
            </form>
            <form method="POST" action="/health/${r.id}/delete" style="display:inline"
                  onsubmit="return confirm('Delete this health record?')">
              <button type="submit" class="btn btn-sm btn-danger">Delete</button>
            </form>
          </td>
        </tr>
      `;
    }).join("");
    body += `</tbody></table></div>`;
  }

  return htmlPage("Health Records", body);
}

// ─── Health Create ────────────────────────────────────────────────────

async function handleHealthCreate(req: Request): Promise<Response> {
  const formData = await req.formData();
  const cattle_id_raw = formData.get("cattle_id")?.toString().trim();
  const date = formData.get("date")?.toString().trim();
  const concern = formData.get("concern")?.toString().trim();
  const resolved = formData.get("resolved") === "1" ? 1 : 0;

  if (!cattle_id_raw || !date || !concern) {
    return handleHealthOverview(true, "Cattle, Date, and Concern are required.", {
      cattle_id: cattle_id_raw || "",
      date: date || "",
      concern: concern || "",
      resolved: resolved ? "1" : "0",
    });
  }

  const cattle_id = parseInt(cattle_id_raw, 10);
  createHealthRecord(cattle_id, date, concern, resolved);
  return redirect("/health");
}

// ─── Health Toggle ────────────────────────────────────────────────────

function handleHealthToggle(id: number): Response {
  toggleHealthRecord(id);
  return redirect("/health");
}

// ─── Health Delete ────────────────────────────────────────────────────

function handleHealthDelete(id: number): Response {
  deleteHealthRecord(id);
  return redirect("/health");
}

// ─── Breeding Overview ────────────────────────────────────────────────

function handleBreedingOverview(showForm: boolean = false, error?: string, defaults?: Record<string, string>): Response {
  const records = getBreedingRecords();

  let body = "";

  // "Record Breeding" button
  body += `
    <div class="list-header">
      <div class="list-header-top">
        <a href="/breeding?form=1" class="btn">+ Record Breeding</a>
        <span class="cattle-count"><strong>${records.length}</strong> breeding record${records.length !== 1 ? "s" : ""}</span>
      </div>
    </div>
  `;

  // Show inline form if requested
  if (showForm) {
    const femaleCount = getFemaleCattle().length;
    if (femaleCount === 0) {
      body += `<div class="alert alert-error">No female cattle recorded yet. <a href="/cattle/add">Add a Cow or Heifer</a> first.</div>`;
    } else {
      body += `
        ${error ? `<div class="alert alert-error">${escapeHTML(error)}</div>` : ""}
        <form method="POST" action="/breeding">
          <h3>Record a Breeding Event</h3>
          <label for="cow_id">Cow *</label>
          <select id="cow_id" name="cow_id" required>
            <option value="">— Select a Cow or Heifer —</option>
            ${femaleCattleOptions(defaults?.cow_id ? parseInt(defaults.cow_id, 10) : null)}
          </select>

          <label for="bull_tag">Bull Tag *</label>
          <input type="text" id="bull_tag" name="bull_tag" required maxlength="50"
                 value="${escapeHTML(defaults?.bull_tag || "")}">

          <label for="breeding_date">Breeding Date *</label>
          <input type="date" id="breeding_date" name="breeding_date" required
                 value="${escapeHTML(defaults?.breeding_date || "")}">

          <label for="notes">Notes</label>
          <textarea id="notes" name="notes" maxlength="500">${escapeHTML(defaults?.notes || "")}</textarea>

          <div class="form-actions">
            <button type="submit" class="btn">Save Breeding Record</button>
            <a href="/breeding" class="btn btn-cancel">Cancel</a>
          </div>
        </form>
      `;
    }
  }

  // Table of records
  if (records.length === 0) {
    body += `<div class="empty-state"><p>No breeding records yet.</p></div>`;
  } else {
    body += `<div class="table-wrapper"><table>`;
    body += `<thead><tr>
      <th>Cow Tag</th><th>Bull Tag</th><th>Breeding Date</th><th>Calf Info</th><th>Actions</th>
    </tr></thead><tbody>`;
    body += records.map(r => {
      let calfInfo: string;
      if (r.calf_tag) {
        const sexEmoji = r.calf_sex === "Heifer" ? "🐮" : "🐂";
        calfInfo = `${sexEmoji} ${escapeHTML(r.calf_tag)} (${escapeHTML(r.calf_sex || "")})`;
      } else {
        calfInfo = `No calf recorded <a href="/breeding/${r.id}/calf" class="btn btn-sm btn-outline">Add Calf</a>`;
      }

      return `
        <tr>
          <td><a href="/cattle/${r.cow_id}" class="tag-link">${escapeHTML(r.cow_tag)}</a></td>
          <td>${escapeHTML(r.bull_tag)}</td>
          <td>${escapeHTML(r.breeding_date)}</td>
          <td>${calfInfo}</td>
          <td>
            <form method="POST" action="/breeding/${r.id}/delete" style="display:inline"
                  onsubmit="return confirm('Delete this breeding record?')">
              <button type="submit" class="btn btn-sm btn-danger">Delete</button>
            </form>
          </td>
        </tr>
      `;
    }).join("");
    body += `</tbody></table></div>`;
  }

  return htmlPage("Breeding Records", body);
}

// ─── Breeding Create ──────────────────────────────────────────────────

async function handleBreedingCreate(req: Request): Promise<Response> {
  const formData = await req.formData();
  const cow_id_raw = formData.get("cow_id")?.toString().trim();
  const bull_tag = formData.get("bull_tag")?.toString().trim();
  const breeding_date = formData.get("breeding_date")?.toString().trim();
  const notes = formData.get("notes")?.toString().trim() || "";

  if (!cow_id_raw || !bull_tag || !breeding_date) {
    return handleBreedingOverview(true, "Cow, Bull Tag, and Breeding Date are required.", {
      cow_id: cow_id_raw || "",
      bull_tag: bull_tag || "",
      breeding_date: breeding_date || "",
      notes,
    });
  }

  const cow_id = parseInt(cow_id_raw, 10);
  createBreedingRecord(cow_id, bull_tag, breeding_date, notes);
  return redirect("/breeding");
}

// ─── Add Calf Form ────────────────────────────────────────────────────

function handleCalfForm(breedingId: number, error?: string, defaults?: Record<string, string>): Response {
  const record = getBreedingRecordById(breedingId);
  if (!record) return notFound();

  if (record.calf_tag) {
    return htmlPage("Calf Already Recorded", `
      <div class="alert alert-error">A calf (${escapeHTML(record.calf_tag)}) has already been recorded for this breeding.</div>
      <p class="mt-1"><a href="/breeding">← Back to Breeding Records</a></p>
    `);
  }

  let body = `<h2>Add Calf for ${escapeHTML(record.cow_tag)} × ${escapeHTML(record.bull_tag)}</h2>`;
  body += `<p class="text-muted mb-1">Breeding date: ${escapeHTML(record.breeding_date)}</p>`;

  if (error) {
    body += `<div class="alert alert-error">${escapeHTML(error)}</div>`;
  }

  body += `
    <form method="POST" action="/breeding/${breedingId}/calf">
      <label for="tag_number">Calf Tag Number *</label>
      <input type="text" id="tag_number" name="tag_number" required maxlength="50"
             value="${escapeHTML(defaults?.tag_number || "")}" autofocus>

      <label for="sex">Sex *</label>
      <select id="sex" name="sex" required>
        <option value="">— Select —</option>
        <option value="Bull"${defaults?.sex === "Bull" ? " selected" : ""}>Bull</option>
        <option value="Heifer"${defaults?.sex === "Heifer" ? " selected" : ""}>Heifer</option>
      </select>

      <label for="birth_date">Birth Date *</label>
      <input type="date" id="birth_date" name="birth_date" required
             value="${escapeHTML(defaults?.birth_date || "")}">

      <label for="notes">Notes</label>
      <textarea id="notes" name="notes" maxlength="500">${escapeHTML(defaults?.notes || "")}</textarea>

      <div class="form-actions">
        <button type="submit" class="btn">Save Calf</button>
        <a href="/breeding" class="btn btn-cancel">Cancel</a>
      </div>
    </form>
  `;

  return htmlPage("Add Calf", body);
}

// ─── Calf Create ──────────────────────────────────────────────────────

async function handleCalfCreate(req: Request, breedingId: number): Promise<Response> {
  const record = getBreedingRecordById(breedingId);
  if (!record) return notFound();

  if (record.calf_tag) {
    return redirect("/breeding");
  }

  const formData = await req.formData();
  const tag_number = formData.get("tag_number")?.toString().trim();
  const sex = formData.get("sex")?.toString().trim();
  const birth_date = formData.get("birth_date")?.toString().trim();
  const notes = formData.get("notes")?.toString().trim() || "";

  if (!tag_number || !sex || !birth_date) {
    return handleCalfForm(breedingId, "Tag number, sex, and birth date are required.", {
      tag_number: tag_number || "",
      sex: sex || "",
      birth_date: birth_date || "",
      notes,
    });
  }

  if (sex !== "Bull" && sex !== "Heifer") {
    return handleCalfForm(breedingId, "Sex must be Bull or Heifer.", {
      tag_number,
      sex: "",
      birth_date,
      notes,
    });
  }

  try {
    createCalf(breedingId, tag_number, sex, birth_date, notes);
  } catch (err: any) {
    if (err.message?.includes("UNIQUE")) {
      return handleCalfForm(breedingId, `A calf with tag "${tag_number}" already exists.`, {
        tag_number: "",
        sex,
        birth_date,
        notes,
      });
    }
    throw err;
  }

  return redirect("/breeding");
}

// ─── Breeding Delete ──────────────────────────────────────────────────

function handleBreedingDelete(id: number): Response {
  deleteBreedingRecord(id);
  return redirect("/breeding");
}

// ─── Export ───────────────────────────────────────────────────────────

function handleExport(): Response {
  let body = "";

  body += `
    <div class="export-page">
      <div class="export-card">
        <h2>📊 Herd Spreadsheet Export</h2>
        <p class="export-description">Download a complete snapshot of your herd as an Excel spreadsheet. The file includes three sheets:</p>
        <ul class="export-sheets">
          <li><strong>Cattle</strong> — every animal with tag, breed, sex, birth date, pasture, health status, and notes</li>
          <li><strong>Breeding Records</strong> — all breeding events with cow, bull, and calf details</li>
          <li><strong>Health Records</strong> — every health concern logged, with active/resolved status</li>
        </ul>
        <a href="/export/download" class="btn btn-download">
          ⬇ Download Herd Spreadsheet
        </a>
        <p class="export-note">File format: <code>.xlsx</code> — opens in Excel, LibreOffice, or Google Sheets</p>
      </div>
    </div>
  `;

  return htmlPage("Export", body);
}

// ─── Export Download ───────────────────────────────────────────────────

async function handleExportDownload(): Promise<Response> {
  const today = new Date().toISOString().split("T")[0];
  const data = getFullExportData();

  const workbook = new ExcelJS.Workbook();

  const headerStyle = {
    font: { bold: true, color: { argb: "FFFFFFFF" } },
    fill: {
      type: "pattern" as const,
      pattern: "solid" as const,
      fgColor: { argb: "FF3D5A1E" },
    },
  };

  const altRowFill = {
    type: "pattern" as const,
    pattern: "solid" as const,
    fgColor: { argb: "FFF0F0F0" },
  };

  function buildSheet(
    name: string,
    headers: string[],
    widths: number[],
    rows: (string | null)[][],
    dateColIndexes: number[] = [],
  ) {
    const sheet = workbook.addWorksheet(name);

    // Set column widths
    sheet.columns = widths.map((w, i) => ({
      header: headers[i],
      key: `col${i}`,
      width: w,
    }));

    // Style the header row (row 1)
    const headerRow = sheet.getRow(1);
    headerRow.font = headerStyle.font;
    headerRow.fill = headerStyle.fill;

    // Freeze header row
    sheet.views = [{ state: "frozen", ySplit: 1 }];

    // Set date format on date columns
    for (const colIdx of dateColIndexes) {
      sheet.getColumn(colIdx).numFmt = "yyyy-mm-dd";
    }

    // Add data rows with alternating shading
    rows.forEach((rowData, i) => {
      const row = sheet.addRow(rowData.map((v) => v ?? ""));
      if (i % 2 === 1) {
        row.fill = altRowFill;
      }
    });

    // Auto-filter on header row
    const lastColLetter = String.fromCharCode(64 + headers.length);
    sheet.autoFilter = {
      from: `A1`,
      to: `${lastColLetter}1`,
    };
  }

  // Sheet 1: Cattle
  buildSheet(
    "Cattle",
    ["Tag Number", "Breed", "Sex", "Birth Date", "Pasture", "Health Status", "Notes"],
    [15, 15, 10, 15, 20, 30, 30],
    data.cattle.map((c) => [
      c.tag_number,
      c.breed || "",
      c.sex,
      c.birth_date,
      c.pasture_name,
      c.health_status,
      c.notes || "",
    ]),
    [4],
  );

  // Sheet 2: Breeding Records
  buildSheet(
    "Breeding Records",
    ["Cow Tag", "Bull Tag", "Breeding Date", "Calf Tag", "Calf Sex", "Calf Birth Date", "Notes"],
    [15, 15, 15, 15, 10, 15, 30],
    data.breeding.map((b) => [
      b.cow_tag,
      b.bull_tag,
      b.breeding_date,
      b.calf_tag,
      b.calf_sex,
      b.calf_birth_date,
      b.notes || "",
    ]),
    [3, 6],
  );

  // Sheet 3: Health Records
  buildSheet(
    "Health Records",
    ["Cattle Tag", "Date", "Concern", "Status"],
    [15, 15, 40, 12],
    data.health.map((h) => [h.tag_number, h.date, h.concern, h.status]),
    [2],
  );

  const buffer = await workbook.xlsx.writeBuffer();

  return new Response(buffer, {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="cattletrackermt-export-${today}.xlsx"`,
    },
  });
}

// ─── Import Upload Page ────────────────────────────────────────────────

function handleImportPage(error?: string): Response {
  let body = `<div class="import-page">`;

  if (error) {
    body += `<div class="alert alert-error">${escapeHTML(error)}</div>`;
  }

  body += `
    <div class="import-card">
      <h2>📥 Bulk Import Cattle</h2>
      <p class="import-description">Upload an Excel (.xlsx) or CSV file with your cattle records. The file should have a header row with column names.</p>

      <div class="import-instructions">
        <h3>Expected Columns</h3>
        <p>The following columns are recognized (flexible matching, case-insensitive):</p>
        <table class="import-columns-table">
          <thead><tr><th>Column</th><th>Aliases</th><th>Required</th></tr></thead>
          <tbody>
            <tr><td><strong>Tag Number</strong></td><td>Tag #, tag_number, tag, ID</td><td class="required-badge">✅ Required</td></tr>
            <tr><td>Breed</td><td>breed</td><td>Optional</td></tr>
            <tr><td>Sex</td><td>Gender</td><td>Optional (Bull/Cow/Steer/Heifer)</td></tr>
            <tr><td>Birth Date</td><td>DOB, Date of Birth, birth_date</td><td>Optional</td></tr>
            <tr><td>Pasture</td><td>Pasture Name, pasture_name</td><td>Optional</td></tr>
            <tr><td>Notes</td><td>Comments, remarks</td><td>Optional</td></tr>
          </tbody>
        </table>
      </div>

      <form method="POST" action="/import/preview" enctype="multipart/form-data" class="import-form">
        <label for="import_file">Choose File (.xlsx or .csv)</label>
        <input type="file" id="import_file" name="file" accept=".xlsx,.csv" required>
        <div class="form-actions">
          <button type="submit" class="btn">🔍 Preview</button>
          <a href="/cattle" class="btn btn-cancel">Cancel</a>
        </div>
      </form>
    </div>
  `;

  body += `</div>`;

  return htmlPage("Import Cattle", body);
}

// ─── Import Preview ───────────────────────────────────────────────────

async function handleImportPreview(req: Request): Promise<Response> {
  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return handleImportPage("Failed to read uploaded file.");
  }

  const file = formData.get("file");
  if (!file || !(file instanceof File)) {
    return handleImportPage("Please select a file to upload.");
  }

  const filename = file.name || "upload.xlsx";
  const lower = filename.toLowerCase();
  if (!lower.endsWith(".xlsx") && !lower.endsWith(".csv")) {
    return handleImportPage("Unsupported file type. Please upload a .xlsx or .csv file.");
  }

  let rows: string[][];
  try {
    const buffer = await file.arrayBuffer();
    rows = await parseFile(buffer, filename);
  } catch (err: any) {
    return handleImportPage(`Failed to parse file: ${err.message || "Unknown error"}`);
  }

  if (rows.length === 0) {
    return handleImportPage("File contains no data. Please check the file and try again.");
  }

  // First row is headers
  const headerRow = rows[0]!;
  const columnMap: Record<number, string> = {};
  for (let i = 0; i < headerRow.length; i++) {
    const field = normalizeHeader(headerRow[i]!);
    if (field) {
      columnMap[i] = field;
    }
  }

  // Check that we have tag_number mapped
  const hasTagColumn = Object.values(columnMap).includes("tag_number");
  if (!hasTagColumn) {
    return handleImportPage("Could not find a 'Tag Number' column. Please include a column labeled Tag Number, Tag #, tag_number, tag, or ID.");
  }

  // Parse data rows
  const dataRows = rows.slice(1).filter(r => r.some(cell => cell.trim() !== "")); // skip empty rows
  const records: ImportRecord[] = [];

  for (const row of dataRows) {
    const record: ImportRecord = {
      tag_number: "",
      breed: "",
      sex: "",
      birth_date: null,
      pasture: "",
      notes: "",
      errors: [],
    };

    for (const [colIdx, field] of Object.entries(columnMap)) {
      const idx = parseInt(colIdx, 10);
      const value = (row[idx] || "").trim();

      switch (field) {
        case "tag_number": record.tag_number = value; break;
        case "breed": record.breed = value; break;
        case "sex": record.sex = normalizeSex(value); break;
        case "birth_date": record.birth_date = value || null; break;
        case "pasture": record.pasture = value; break;
        case "notes": record.notes = value; break;
      }
    }

    // Validate
    record.errors = validateRecord(record);

    // Check for duplicate tags in the database
    if (record.tag_number && getCattleByTag(record.tag_number)) {
      record.errors.push(`Tag "${record.tag_number}" already exists in the database`);
    }

    // Check for duplicate tags within the file itself
    const duplicateInPreview = records.filter(r => r.tag_number === record.tag_number).length;
    if (record.tag_number && duplicateInPreview > 0) {
      record.errors.push(`Tag "${record.tag_number}" appears more than once in this file`);
    }

    records.push(record);
  }

  // Store in memory for the import step
  previewData = records;

  // Build HTML preview
  let body = `<div class="import-page">`;
  body += `<div class="import-card">`;
  body += `<h2>📋 Preview Import</h2>`;
  body += `<p class="import-description">Found <strong>${records.length}</strong> record${records.length !== 1 ? "s" : ""} in <em>${escapeHTML(filename)}</em>.</p>`;

  const validCount = records.filter(r => r.errors.length === 0).length;
  const errorCount = records.length - validCount;

  if (errorCount > 0) {
    body += `<div class="alert alert-error">⚠️ ${errorCount} row${errorCount !== 1 ? "s have" : " has"} validation issues — rows with errors will be skipped.</div>`;
  }

  if (validCount === 0) {
    body += `<div class="alert alert-error">No valid records to import. All rows have errors.</div>`;
    body += `<div class="form-actions"><a href="/import" class="btn">← Try Again</a></div>`;
  } else {
    body += `
      <form method="POST" action="/import" class="import-form">
        <p class="mb-1">Click "Import All" to import <strong>${validCount}</strong> valid record${validCount !== 1 ? "s" : ""}.</p>
        <div class="form-actions">
          <button type="submit" class="btn btn-download">✅ Import ${validCount} Cattle</button>
          <a href="/import" class="btn btn-cancel">Cancel</a>
        </div>
      </form>
    `;
  }

  // Preview table — show first 10 rows
  const previewRows = records.slice(0, 10);
  body += `<h3 class="mt-1">Preview (first ${previewRows.length} row${previewRows.length !== 1 ? "s" : ""})</h3>`;
  body += `<div class="table-wrapper"><table class="import-preview-table">`;
  body += `<thead><tr>
    <th>#</th><th>Tag Number</th><th>Breed</th><th>Sex</th><th>Birth Date</th><th>Pasture</th><th>Notes</th><th>Status</th>
  </tr></thead><tbody>`;

  body += previewRows.map((r, i) => {
    const status = r.errors.length === 0
      ? `<span class="health-ok">✅ OK</span>`
      : `<span class="health-active">❌ ${escapeHTML(r.errors.join("; "))}</span>`;
    return `
      <tr class="${r.errors.length > 0 ? "import-row-error" : ""}">
        <td>${i + 1}</td>
        <td><strong>${escapeHTML(r.tag_number || "—")}</strong></td>
        <td>${escapeHTML(r.breed || "—")}</td>
        <td>${escapeHTML(r.sex || "—")}</td>
        <td>${escapeHTML(r.birth_date || "—")}</td>
        <td>${escapeHTML(r.pasture || "—")}</td>
        <td>${escapeHTML(r.notes || "—")}</td>
        <td>${status}</td>
      </tr>
    `;
  }).join("");

  body += `</tbody></table></div>`;

  if (records.length > 10) {
    body += `<p class="text-muted">… and ${records.length - 10} more row${records.length - 10 !== 1 ? "s" : ""} not shown.</p>`;
  }

  body += `</div></div>`;

  return htmlPage("Preview Import", body);
}

// ─── Import Execute ───────────────────────────────────────────────────

function handleImportExecute(): Response {
  if (previewData.length === 0) {
    return redirect("/import");
  }

  const validRecords = previewData.filter(r => r.errors.length === 0);
  const skipped = previewData.filter(r => r.errors.length > 0);

  // Resolve pasture names to IDs
  const pastures = db.query("SELECT id, name FROM pastures").all() as { id: number; name: string }[];
  const pastureMap = new Map<string, number>();
  for (const p of pastures) {
    pastureMap.set(p.name.toLowerCase(), p.id);
  }

  const toImport = validRecords.map(r => {
    let pasture_id: number | null = null;
    if (r.pasture) {
      pasture_id = pastureMap.get(r.pasture.toLowerCase()) ?? null;
    }
    return {
      tag_number: r.tag_number,
      breed: r.breed || "",
      sex: r.sex || "",
      birth_date: r.birth_date,
      pasture_id,
      notes: r.notes || "",
    };
  });

  const result = importCattle(toImport);

  // Collect all skipped reasons
  const allSkipped = [
    ...skipped.map(r => ({ tag: r.tag_number || "(missing tag)", reason: r.errors.join("; ") })),
    ...result.skipped,
  ];

  const imported = result.imported;

  // Redirect to /cattle with summary
  const params = new URLSearchParams();
  params.set("imported", String(imported));
  params.set("skipped", String(allSkipped.length));

  // Pass skipped details via a comma-encoded format (simple approach)
  if (allSkipped.length > 0) {
    const skippedDetails = allSkipped.map(s => `${s.tag}: ${s.reason}`).join("||");
    params.set("skipped_details", skippedDetails);
  }

  // Clear preview data
  previewData = [];

  return redirect(`/cattle?${params.toString()}`);
}

// ─── Server ───────────────────────────────────────────────────────────

const server = Bun.serve({
  port: 3000,
  async fetch(req) {
    const url = new URL(req.url);
    const pathname = url.pathname;
    const method = req.method;

    // Serve static files
    if (pathname.startsWith("/public/") || pathname === "/style.css") {
      const staticPath = pathname.replace(/^\/public/, "");
      return serveStatic(staticPath) ?? new Response("Not Found", { status: 404 });
    }

    // GET routes
    if (method === "GET") {
      if (pathname === "/") return redirect("/cattle");
      if (pathname === "/cattle") {
        const imported = url.searchParams.get("imported");
        const skipped = url.searchParams.get("skipped");
        const importSummary = imported ? {
          imported: parseInt(imported, 10),
          skipped: parseInt(skipped || "0", 10),
          skippedDetails: url.searchParams.get("skipped_details") || undefined,
        } : undefined;
        return handleCattleList(
          url.searchParams.get("search") || undefined,
          url.searchParams.get("added"),
          importSummary
        );
      }
      if (pathname === "/cattle/add") return handleCattleAddForm();
      if (pathname === "/pastures") return handlePasturesList();
      if (pathname === "/health") return handleHealthOverview(
        url.searchParams.get("form") === "1",
      );
      if (pathname === "/breeding") return handleBreedingOverview(
        url.searchParams.get("form") === "1",
      );
      if (pathname === "/export") return handleExport();
      if (pathname === "/export/download") return await handleExportDownload();
      if (pathname === "/import") return handleImportPage();

      // Dynamic GET routes
      const cattleDetailMatch = pathname.match(/^\/cattle\/(\d+)$/);
      if (cattleDetailMatch) return handleCattleDetail(parseInt(cattleDetailMatch[1]!, 10));

      const cattleEditMatch = pathname.match(/^\/cattle\/(\d+)\/edit$/);
      if (cattleEditMatch) return handleCattleEditForm(parseInt(cattleEditMatch[1]!, 10));

      const pastureDetailMatch = pathname.match(/^\/pastures\/(\d+)$/);
      if (pastureDetailMatch) return handlePastureDetail(
        parseInt(pastureDetailMatch[1]!, 10),
        url.searchParams.get("search") || undefined
      );

      const calfFormMatch = pathname.match(/^\/breeding\/(\d+)\/calf$/);
      if (calfFormMatch) return handleCalfForm(parseInt(calfFormMatch[1]!, 10));
    }

    // POST routes
    if (method === "POST") {
      if (pathname === "/cattle") return handleCattleCreate(req);
      if (pathname === "/pastures") return handlePastureCreate(req);
      if (pathname === "/breeding") return handleBreedingCreate(req);
      if (pathname === "/health") return handleHealthCreate(req);

      if (pathname === "/import/preview") return await handleImportPreview(req);
      if (pathname === "/import") return handleImportExecute();

      const cattleUpdateMatch = pathname.match(/^\/cattle\/(\d+)$/);
      if (cattleUpdateMatch) return handleCattleUpdate(req, parseInt(cattleUpdateMatch[1]!, 10));

      const cattleDeleteMatch = pathname.match(/^\/cattle\/(\d+)\/delete$/);
      if (cattleDeleteMatch) return handleCattleDelete(parseInt(cattleDeleteMatch[1]!, 10));

      const pastureDeleteMatch = pathname.match(/^\/pastures\/(\d+)\/delete$/);
      if (pastureDeleteMatch) return handlePastureDelete(parseInt(pastureDeleteMatch[1]!, 10));

      const calfCreateMatch = pathname.match(/^\/breeding\/(\d+)\/calf$/);
      if (calfCreateMatch) return handleCalfCreate(req, parseInt(calfCreateMatch[1]!, 10));

      const breedingDeleteMatch = pathname.match(/^\/breeding\/(\d+)\/delete$/);
      if (breedingDeleteMatch) return handleBreedingDelete(parseInt(breedingDeleteMatch[1]!, 10));

      const healthToggleMatch = pathname.match(/^\/health\/(\d+)\/toggle$/);
      if (healthToggleMatch) return handleHealthToggle(parseInt(healthToggleMatch[1]!, 10));

      const healthDeleteMatch = pathname.match(/^\/health\/(\d+)\/delete$/);
      if (healthDeleteMatch) return handleHealthDelete(parseInt(healthDeleteMatch[1]!, 10));
    }

    return notFound();
  },
});

console.log(`🏔️ CattleTrackerMt running at http://localhost:${server.port}`);
