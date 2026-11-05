# EduTask Performance V14 — Installation & Validation

This is a *compatible replacement* for EduTask V13.3 `main.html`, V6 `config.js` API client, and V13.2 `Code.gs` (which already includes the V11 mentorship modules and V10 smart operations). No Sheets recreation, account bootstrap, or migration is required. Back up your files and Google Sheet before production deployment.

## What actually changes

- **Startup:** Display anonymous app shell and Finora-style loading skeleton first. Verify the current session via `auth.session` (small response), *then* unlock protected menus. Fetch `dashboard.mentor`/`dashboard.intern` in the background. Cached numeric aggregates are only rendered after successful session validation. The initial dashboard may still take time if the server is cold.
- **Navigation:** Keep the current route title and skeleton immediately, use per-route in-memory caching, avoid repeating `users.list` while usable user references are already loaded, and prefetch form references after the dashboard rather than blocking the form on each opening.
- **Task details:** Show the known task title/status/progress immediately from the already-loaded list. Fetch comments, submissions and reviews separately. On first open, details still need a real backend response.
- **Backend:** Add an ephemeral per-token authorization cache for six seconds, with explicit invalidation after logout/password reset/role change/deactivation. Dashboard data cached by account ID + role + global change version for 22 seconds (no cross-account cached projection). Add a similarly short summary-metrics cache.
- **Email and OTP:** Reduce one Sheets write for successful OTP validation; release the global script lock before calling MailApp for invitations, registration and password resets. Email transport latency is still external.
- **Cross-account sync:** Check `sync.version` approximately every 4.5 seconds in visible active tabs, more slowly in hidden tabs. Read the active module again only after the change marker changes; the pending editor's form is not overwritten. Same-origin tabs can use BroadcastChannel to notify promptly. This is **not guaranteed real-time within 1 second across devices**.
- **Diagnostics:** `EdutaskPerf.report()` prints API measurements with p50 and p95; `EdutaskSpeed.report()` reports route cache state and sync polling interval.

## Install

1. Back up current `main.html`, `config.js`, `Code.gs`, and the entire Google Sheets database.
2. Open the existing Apps Script project, replace **only** `Code.gs` with V14, Save. Do not run `bootstrapInitialAdmin()` or create another spreadsheet.
3. Apps Script: **Deploy → Manage deployments → Edit current deployment → New version → Deploy**. Keep the same `/exec` URL.
4. In your VS Code EduTask folder, replace `main.html` and `config.js`. Keep `api/proxy.js`, `index.html` and `verify.html` as they are if your Fast Auth V3 credentials already work.
5. In Vercel, ensure `EDUTASK_GAS_URL`, `EDUTASK_KDF_PEPPER`, and `EDUTASK_PROXY_SIGNING_KEY` remain set as appropriate for Production and Development. Never add these secrets to HTML or Git.
6. Start local using `npx vercel@latest pull --environment=development` and `npx vercel@latest dev`. Visit the local URL printed by the CLI.
7. Log in as admin, mentor, and intern in separate browser profiles. Test login, session restoration, navigation, form submission, task detail, OTP, password reset, logout, and cross-account updates with test data.
8. Only after successful local checks: `git add main.html config.js && git commit -m "EduTask V14 performance" && git push origin main`.

## Real-world performance tests

Open Chrome DevTools → Console:

```js
EdutaskPerf.report();
EdutaskSpeed.report();
EdutaskLive.status();
```

Then test with DevTools → Network (`Disable cache` OFF for warm tests), refresh twice and observe both **time to shell** and **time to verified data**. Use a normal network condition. For each API request compare `roundtripMs` and `backendMs` before/after V14. Cold Apps Script start, large Sheets, proxy network hops and MailApp may still exceed 1–3 seconds. The simulated millisecond readings in the code test are not production benchmarks.

### Why Sheets API v4 isn't the main lever

The installed backend uses `SpreadsheetApp` inside Apps Script; it does **not** simply become faster by renaming it Sheets API v4. Google's recommendation is to minimize service calls, batch reads/writes, and use caches. `CacheService` cannot replace a robust low-latency database at larger scale.

### One-second synchronization limit

With Apps Script + Google Sheets + periodic polling, a strict **≤1 second across admin/mentor/intern devices is not guaranteed**. To support true near-real-time push, add a dedicated realtime transport (e.g. Firebase Realtime Database, Firestore listeners, or Supabase Realtime) with tenant-aware permissions and a backend write confirmation flow. Do not reduce polling to 1 second for every user; it can exhaust Apps Script execution and concurrent request quotas.

## Testing performed

- JavaScript syntax checks for inline `main.html`, `config.js`, and `Code.gs` succeeded.
- Fake-Google-Sheets backend test succeeded for setup of 44 tabs, bootstrap admin, login, authorized sessions, cache hit, dashboard, references, invite, registration, email verification, role-cache invalidation, password reset/session revocation, and logout.
- Automated headless browser navigation was attempted but blocked by the execution environment. Desktop/mobile visual interactions must be tested in your local Vercel environment.
- This package has **not** been connected to your actual Vercel project, Google account or production Google Sheets.
