export interface PageUser {
  email: string;
  trialDaysLeft?: number | null;
  isSubscribed?: boolean;
}

export function renderPage(title: string, bodyHTML: string, user?: PageUser | null): string {
  // Compute trial badge for the navbar
  let trialBadge = "";
  if (user && !user.isSubscribed) {
    if (user.trialDaysLeft !== null && user.trialDaysLeft !== undefined && user.trialDaysLeft > 0) {
      trialBadge = `<span class="trial-badge trial-badge-active">${user.trialDaysLeft} day${user.trialDaysLeft !== 1 ? "s" : ""} left</span>`;
    } else {
      trialBadge = `<span class="trial-badge trial-badge-expired">Trial expired</span>`;
    }
  }

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHTML(title)} — CattleTrackerMt</title>
  <link rel="stylesheet" href="/style.css">
</head>
<body>
  <nav class="navbar">
    <a href="/" class="brand">🏔️ CattleTrackerMt</a>
    <div class="nav-links">
      ${user ? `
        <a href="/cattle">Cattle</a>
        <a href="/pastures">Pastures</a>
        <a href="/breeding">Breeding</a>
        <a href="/health">Health</a>
        <a href="/import">Import</a>
        <a href="/scan">Scan</a>
        <a href="/export">Export</a>
        <span class="nav-user">${escapeHTML(user.email)}</span>
        ${trialBadge}
        <form method="POST" action="/logout" class="nav-logout-form">
          <button type="submit" class="nav-logout-btn">Logout</button>
        </form>
      ` : `
        <a href="/login">Login</a>
        <a href="/register">Register</a>
      `}
    </div>
  </nav>
  <main class="container">
    <h1>${escapeHTML(title)}</h1>
    ${bodyHTML}
  </main>
</body>
</html>`;
}

export function escapeHTML(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
