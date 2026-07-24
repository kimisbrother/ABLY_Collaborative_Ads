// Intentionally does not cache or intercept anything — the dashboard always
// needs a fresh fetch of the Google Sheets CSVs. This exists only so the
// browser considers the site installable as a home-screen app.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", () => {});
