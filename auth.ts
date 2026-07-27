import { db } from "./db";

const SESSION_COOKIE = "session";

export interface User {
  id: number;
  email: string;
  created_at: string;
  trial_ends_at: string | null;
  subscribed: number;
}

// ─── Password helpers ──────────────────────────────────────────────────

export async function hashPassword(password: string): Promise<string> {
  return Bun.password.hash(password, { algorithm: "bcrypt", cost: 10 });
}

export async function verifyPassword(password: string, hash: string): Promise<boolean> {
  return Bun.password.verify(password, hash);
}

// ─── Session helpers ───────────────────────────────────────────────────

export function createSession(userId: number): { token: string; cookie: string } {
  const token = crypto.randomUUID();
  db.run(
    `INSERT INTO sessions (token, user_id) VALUES (?, ?)`,
    [token, userId]
  );
  const cookie = `${SESSION_COOKIE}=${token}; HttpOnly; SameSite=Lax; Secure; Path=/; Max-Age=${60 * 60 * 24 * 30}`;
  return { token, cookie };
}

export function getUserFromSession(req: Request): User | null {
  const cookieHeader = req.headers.get("cookie");
  if (!cookieHeader) return null;

  const cookies = parseCookies(cookieHeader);
  const token = cookies[SESSION_COOKIE];
  if (!token) return null;

  const row = db.query(`
    SELECT u.id, u.email, u.created_at, u.trial_ends_at, u.subscribed
    FROM sessions s
    JOIN users u ON u.id = s.user_id
    WHERE s.token = ?
  `).get(token) as User | null;

  return row ?? null;
}

export function deleteSession(req: Request): string | null {
  const cookieHeader = req.headers.get("cookie");
  if (!cookieHeader) return null;

  const cookies = parseCookies(cookieHeader);
  const token = cookies[SESSION_COOKIE];
  if (!token) return null;

  db.run(`DELETE FROM sessions WHERE token = ?`, [token]);

  // Return a cookie-clearing header
  return `${SESSION_COOKIE}=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0`;
}

// ─── Route guards ──────────────────────────────────────────────────────

export async function requireAuth(req: Request): Promise<User> {
  const user = getUserFromSession(req);
  if (!user) {
    // We throw a special value that the router catches to redirect
    throw new AuthRequiredError(req);
  }
  return user;
}

export async function requireGuest(req: Request): Promise<User | null> {
  const user = getUserFromSession(req);
  if (user) {
    throw new AlreadyLoggedInError();
  }
  return null;
}

// ─── Errors ────────────────────────────────────────────────────────────

export class AuthRequiredError extends Error {
  redirectTo: string;
  constructor(req: Request) {
    super("Authentication required");
    const url = new URL(req.url);
    this.redirectTo = `/login?redirect=${encodeURIComponent(url.pathname + url.search)}`;
  }
}

export class AlreadyLoggedInError extends Error {
  constructor() {
    super("Already logged in");
  }
}

// ─── Cookie parser ─────────────────────────────────────────────────────

function parseCookies(header: string): Record<string, string> {
  const cookies: Record<string, string> = {};
  for (const part of header.split(";")) {
    const eq = part.indexOf("=");
    if (eq === -1) continue;
    const key = part.slice(0, eq).trim();
    const val = part.slice(eq + 1).trim();
    cookies[key] = val;
  }
  return cookies;
}
