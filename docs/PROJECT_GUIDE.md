# ACE Content Portal — Project and Deployment Guide

## 1. Architecture

ACE Content Portal is a Google Apps Script Web App with two runtime source files:

- `code.gs` implements `doGet`, login and dashboard APIs, KPI processing, User Mapping, attendance integration, announcements, access control, logging, and avatar cache operations.
- `index.html` contains the login page, responsive sidebar, admin and VJ/MC views, CSS, and browser-side JavaScript. It calls backend functions through `google.script.run`.

The project uses existing Google Sheets as its data stores. It also reads the external attendance workbook and uses Drive for cached profile images. The source contains the configured spreadsheet IDs; keep them aligned with the intended environment.

## 2. Deploy to a Google Apps Script project

1. Open the Apps Script project that is bound to or authorized for the ACE workbook.
2. Create or update the script file `code.gs` with the repository version.
3. Create an HTML file named `index` (Apps Script adds the `.html` extension) and paste the contents of `index.html`.
4. Save all files. Check the Apps Script editor for parse errors before opening the Web App.
5. In the editor, run the required setup/authorization functions when the project asks for Google Sheets, Drive, or URL Fetch permissions. Review the requested permissions and authorize using the intended owner account.
6. Select **Deploy → Test deployments → Web app**, copy the `/dev` URL, and test it while signed in as a project Editor.
7. After changing source, save the project and reload `/dev` to test the latest saved code.

## 3. Data and schema expectations

- Do not remove or rename existing sheets or columns relied on by the current project.
- Preserve columns A–J in `User_Mapping`.
- The project may create supporting sheets for logs, announcements, snapshots, avatar references, or local attendance corrections when required by existing code.
- Attendance source data is expected in an `Attendance_Logs` sheet. Shift rules may be read from `Shift_Settings` when available.
- Sheet names, header names, and access rights must match the connected workbooks. Verify permissions using the account that executes the Web App.

## 4. Access control and production readiness

- Review the Admin and VJ/MC PIN fallback branches in `code.gs` before exposing a deployment to a broad audience. Replace test/default credentials with organization-managed authentication before production use.
- Review `checkAccessPermission`, the user/IP lock behavior, and Logger tracking after any authentication change.
- Treat spreadsheet IDs and profile data as internal configuration. Do not commit OAuth tokens, API keys, real credentials, or exported user data.
- A test deployment ending in `/dev` is intended for development and is restricted to users who can edit the Apps Script project.
- Use an appropriately restricted versioned deployment for production and verify “Execute as” and “Who has access” settings with the workspace administrator.

## 5. Troubleshooting

### Loading overlay stays visible

Check DevTools Console for the first syntax error. In the Apps Script editor, verify that there is only one `withSuccessHandler` in the `loadMonthList()` chain and that the frontend calls a backend function present in `code.gs`. Save both files before reloading `/dev`.

### Spreadsheet or authorization errors

Confirm that the executing account can access each configured workbook and that Apps Script has the required scopes. Run the relevant function from the editor to trigger the Google authorization flow, then retry the Web App.

### Attendance data does not appear

Confirm the external workbook has the expected `Attendance_Logs` sheet and header fields for name, date, time in, and time out. Check the Apps Script execution log for access or header errors.

### External resources are blocked

The page references third-party CDN assets and external services. Browser extensions, network policy, or content blockers can block those requests. The CSP reporting `ERR_BLOCKED_BY_CLIENT` message may come from a browser blocker; diagnose it separately from application JavaScript errors.

## 6. Source synchronization

The GitHub repository is the source handoff location; it is not automatically connected to Apps Script. Copy each changed file into the matching Apps Script file and save it before testing. Keep the repository and Apps Script project synchronized so `/dev` runs the code represented in Git.
