# EduTask — Smarter Tasks. Better Growth.

A source-code implementation for an internal internship workspace with 5 primary source files:
`index.html`, `verify.html`, `main.html`, `config.js`, `Code.gs`.

**Mandatory sixth infrastructure file:** `api/proxy.js`. Google Apps Script ContentService
cannot provide arbitrary CORS headers/preflight behavior for a Vercel-hosted app.
Vercel's server-side function proxies POST requests to GAS, follows Google's redirects,
and keeps the GAS deployment URL in an environment variable instead of frontend code.

## Deployment

1. Create a Google Spreadsheet. Copy its ID from `docs.google.com/spreadsheets/d/ID/edit`.
2. Create a standalone Apps Script project. Paste `Code.gs`. Enable V8.
3. Under **Project Settings → Script properties**, set:
   - `EDUTASK_SPREADSHEET_ID` = Spreadsheet ID
   - `EDUTASK_APP_URL` = Vercel frontend origin, e.g. `https://your-project.vercel.app`
   - `EDUTASK_BOOTSTRAP_EMAIL` = administrator's email
   - `EDUTASK_BOOTSTRAP_PASSWORD` = unique temporary password (16+ characters)
4. In Apps Script, run `setupEdutask()` once, authorize Sheets and MailApp. Verify 28 sheets.
5. Run `bootstrapInitialAdmin()` once. Password bootstrap property is removed automatically.
   Sign in with your bootstrap email and password, then change that password.
6. Deploy Apps Script: **Deploy → New deployment → Web app**. Execute as **Me**, access **Anyone**.
   This makes the API URL publicly reachable, but API data still requires server-side
   authentication. Do not publish the spreadsheet itself.
7. Deploy this entire repository to Vercel using the static project preset. Do not
   deploy only the five main files: include the mandatory `api/proxy.js`.
8. Set the Vercel environment variable `EDUTASK_GAS_URL` to the Apps Script Web App
   `/exec` URL. Redeploy to pick up the environment variable.
9. Open `/index.html`. Log in as initial admin. Invite interns/mentors from Intern
   Management. Registrants use the emailed invitation code, verify email, then log in.
10. After editing `Code.gs`, create a **new deployment version** or update existing
    deployment to the newest version. Test API from Vercel after every change.

## Test sequence

- Admin signs in → Invite intern → Intern registers → OTP verification → Intern signs in.
- Admin creates project and assigns a task to the intern.
- Intern sees task, updates progress, adds comments, submits a URL or note.
- Admin/mentor reviews submission, requests revision, intern resubmits, reviewer approves.
- Admin exports CSV, then signs out; ensure an invalid session is rejected.

## Features actually implemented in source

- Invitation-only registration, OTP email verification and resend, password reset,
  salted PBKDF2-HMAC-SHA256 password hashing, secret pepper, token-hashed sessions,
  backend session validation and revocation, role/access checks, email throttling.
- CRUD/API for projects, milestones, batches, teams, programs, templates,
  announcements, task/subtask/comments, daily reports, intern profiles,
  project members, mentor assignments, evaluation criteria, and app settings.
- Individual task assignments, progress, submissions, review/revisions/approval,
  notifications, activity logs, task analytics and CSV exports.
- Two role-sensitive SPA workspaces, mobile layout, interactive Kanban/table/calendar,
  profile/password, and reports with backend-owned data; 28 sheet setup.

## Important constraints and security review

**Do not assume this project is production-security certified.** Google Apps Script +
Sheets have execution, concurrency, MailApp, and database-size constraints. The
custom credential implementation requires independent security review, controlled
load tests, monitoring, backup/restoration, penetration testing, and operations policies.

- PBKDF2 is intentionally low iteration (`12000`) to fit Apps Script's Utilities HMAC
  costs and a serverless proxy's request duration. This does not meet modern native
  PBKDF2 cost recommendations. For sensitive production user information, use an
  established identity provider or a server-side platform with native password hashing.
- `Utilities.getUuid()` entropy is not documented as a secure random source;
  token/OTP generation should be replaced with an audited CSPRNG for high-security use.
- Session tokens live in `sessionStorage`, which means protecting against XSS is
  essential and sessions end when the tab/browser session closes. No Remember Me.
- Email notifications depend on MailApp daily quotas. Avoid high-volume invitations.
- Google Sheets has no actual foreign keys or transactions. Locks limit races but
  do not provide relational transactional guarantees across tables.
- There is no binary upload in this version: use links to access-controlled files.
- Client synchronization uses refresh/on-focus and 120-second polling; not WebSockets.
- CSV export is provided; PDF export is not implemented.
- No tests have been run against your real Google account, Apps Script deployment, or
  Vercel project. End-to-end acceptance criteria require deployment and manual QA.
- Some advanced capabilities from the specification are not complete, including
  automated schedules/reminders, bulk workflows, rich permission matrix editor,
  exhaustive archive restoration, sophisticated analytics slicing, and attachment
  storage. This package implements a functional core with extensible records/modules,
  not a claim that all 1,659 lines of requested specifications are fully satisfied.

## API shape

`POST /api/proxy` JSON `{action, payload, token, requestId?}`.
Response `{success, data, code, message, timestamp}`. The proxy sends requests to
Google Apps Script and follows its redirects.

## Operational performance

A fast static shell and instant client-side navigation are targeted, not a guarantee.
Real Apps Script request duration is network/deployment dependent. Measure LCP, INP,
CLS, API duration and Sheets row growth with Web Vitals and Apps Script execution logs.
For larger teams, use an API backend and a real database rather than Sheets.
