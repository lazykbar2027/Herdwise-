export function renderPage(title: string, bodyHTML: string): string {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHTML(title)} — HerdWise</title>
  <link rel="stylesheet" href="/style.css">
</head>
<body>
  <nav class="navbar">
    <a href="/" class="brand">🐄 HerdWise</a>
    <div class="nav-links">
      <a href="/cattle">Cattle</a>
      <a href="/pastures">Pastures</a>
      <a href="/breeding">Breeding</a>
      <a href="/health">Health</a>
      <a href="/import">Import</a>
      <a href="/export">Export</a>
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
