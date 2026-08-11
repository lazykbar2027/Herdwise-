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

// ─── Public marketing landing page (no auth required) ────────────────

const LANDING_FEATURES: Array<{ icon: string; title: string; blurb: string }> = [
  {
    icon: "🐄",
    title: "Manage Your Herd",
    blurb: "Add cattle with tag numbers, breeds, and assign to pastures. Search and filter in seconds.",
  },
  {
    icon: "🌿",
    title: "Track Pastures",
    blurb: "Know exactly which animals are in which pasture at a glance.",
  },
  {
    icon: "🐂",
    title: "Breeding & Calving",
    blurb: "Record which bull bred which cow, track calves by sex and birth date.",
  },
  {
    icon: "🩺",
    title: "Health Records",
    blurb: "Log health concerns, mark them resolved, keep a complete history.",
  },
  {
    icon: "📊",
    title: "Excel Export",
    blurb: "One-click download of your entire herd as a formatted spreadsheet.",
  },
  {
    icon: "📱",
    title: "Farm-Ready Mobile",
    blurb: "Works on any phone or tablet. No app to install.",
  },
  {
    icon: "📡",
    title: "EID Tag Scanning",
    blurb: "Use your Bluetooth RFID wand to instantly pull up any animal.",
  },
  {
    icon: "📥",
    title: "Bulk Import",
    blurb: "Upload your existing cattle spreadsheet and get started in minutes.",
  },
];

export function renderLandingPage(user?: PageUser | null): string {
  // Compute trial badge for the navbar (same as app pages)
  let trialBadge = "";
  if (user && !user.isSubscribed) {
    if (user.trialDaysLeft !== null && user.trialDaysLeft !== undefined && user.trialDaysLeft > 0) {
      trialBadge = `<span class="trial-badge trial-badge-active">${user.trialDaysLeft} day${user.trialDaysLeft !== 1 ? "s" : ""} left</span>`;
    } else {
      trialBadge = `<span class="trial-badge trial-badge-expired">Trial expired</span>`;
    }
  }

  // Hero CTA: logged-in users go straight to their herd; visitors start a free trial
  const heroCTA = user
    ? `<a class="btn landing-btn" href="/cattle">Go to My Herd →</a>`
    : `<a class="btn landing-btn" href="/register">Start Free Trial</a>`;
  const navRight = user ? `
        <a href="/cattle">Cattle</a>
        <a href="/pastures">Pastures</a>
        <a href="/breeding">Breeding</a>
        <a href="/health">Health</a>
        <span class="nav-user">${escapeHTML(user.email)}</span>
        ${trialBadge}
        <form method="POST" action="/logout" class="nav-logout-form">
          <button type="submit" class="nav-logout-btn">Logout</button>
        </form>
      ` : `
        <a href="/login">Login</a>
        <a href="/register">Register</a>
      `;

  const featureCards = LANDING_FEATURES.map(f =>
    `<div class="landing-card">
       <div class="landing-card-icon">${f.icon}</div>
       <h3>${escapeHTML(f.title)}</h3>
       <p>${escapeHTML(f.blurb)}</p>
     </div>`
  ).join("");

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CattleTrackerMt — Know your herd. Anytime, anywhere.</title>
  <meta name="description" content="Track every animal, pasture, breeding, and health concern — right from your phone or computer.">
  <link rel="stylesheet" href="/style.css">
</head>
<body>
  <nav class="navbar">
    <a href="/" class="brand">🏔️ CattleTrackerMt</a>
    <div class="nav-links">
      ${navRight}
    </div>
  </nav>

  <header class="landing-hero">
    <div class="landing-hero-inner">
      <h1>Know your herd. Anytime, anywhere.</h1>
      <p class="landing-sub">Track every animal, pasture, breeding, and health concern — right from your phone or computer.</p>
      <div class="landing-cta">
        ${heroCTA}
        <a class="btn btn-outline landing-btn" href="#features">Learn More</a>
      </div>
    </div>
  </header>

  <section id="features" class="landing-section">
    <h2 class="landing-section-title">Everything your ranch needs</h2>
    <div class="landing-grid">
      ${featureCards}
    </div>
  </section>

  <section class="landing-section landing-pricing-section">
    <h2 class="landing-section-title">Simple, honest pricing</h2>
    <div class="landing-price-card">
      <p class="landing-price">14-day free trial. Then <strong>$15/month</strong>.</p>
      <ul class="landing-price-list">
        <li>✓ Unlimited cattle</li>
        <li>✓ All features</li>
        <li>✓ Excel export</li>
        <li>✓ Mobile access</li>
      </ul>
      ${user
        ? `<a class="btn landing-btn" href="/cattle">Go to My Herd →</a>`
        : `<a class="btn landing-btn" href="/register">Start Free Trial</a>`}
      <p class="landing-note">No contracts. Cancel anytime.</p>
    </div>
  </section>

  <footer class="landing-footer">
    <p>© 2026 CattleTrackerMt</p>
  </footer>
</body>
</html>`;
}
