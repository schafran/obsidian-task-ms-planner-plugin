# Obsidian <-> Microsoft To Do task sync — design

## Goal

Bidirectional sync between Obsidian tasks (Tasks-plugin markdown syntax) and a dedicated Microsoft To Do list, so due-date tasks get real nudging (To Do/Outlook reminders) instead of getting lost in the vault.

## Scope (v1)

**In scope:**
- Only tasks with a due date (`📅 YYYY-MM-DD`) sync. Undated tasks are untouched.
- Obsidian -> To Do: new due-date task created locally is pushed as a new To Do task (due date carried over).
- To Do -> Obsidian: completing a To Do task marks the Obsidian checkbox done (and sets `✅ YYYY-MM-DD`).
- Obsidian -> To Do: marking the Obsidian checkbox done completes the To Do task.
- Edits (title, due date) sync both ways, last-write-wins on conflict.
- To Do -> Obsidian: a task created directly in To Do (not yet linked to any Obsidian line) is appended to the current ISO-week weekly note's "Other Action Items" section. If that note doesn't exist yet, the task is held in a pending queue and flushed in once the note exists.
- Desktop only (`isDesktopOnly: true`).

**Out of scope (v1):**
- Recurring tasks (`🔁`). Detected and skipped/logged, not synced.
- Tasks without a due date.
- Multiple To Do lists / multiple Obsidian vaults.
- Mobile support (see Token storage — `app.secretStorage` mobile behaviour is unconfirmed; revisit later).
- Subtasks, checklist items, attachments, categories.

## Target service: Microsoft To Do (not Planner)

Personal task list, no Microsoft 365 Group required (unlike Planner, which needs a group-backed or beta-only roster plan). Single delegated scope, simplest fit for a single-user desktop plugin.

- Dedicated To Do list, created by the plugin on first run (or picked from existing lists in settings) — name: "Obsidian".
- Graph endpoints: `GET/POST /me/todo/lists`, `GET/POST/PATCH/DELETE /me/todo/lists/{listId}/tasks`.
- Delegated scope: `Tasks.ReadWrite`. Does not require admin consent in a normal tenant (flag: some tenants restrict user consent org-wide — verify against target tenant during auth setup).
- Relevant `todoTask` fields:
  - `title`
  - `status`: `notStarted` / `completed` (plugin only ever writes these two; `inProgress`/`waitingOnOthers`/`deferred` are read-only inputs from the To Do side — if encountered, treat as `notStarted` for sync purposes)
  - `dueDateTime`: `{ dateTime, timeZone }` — `dateTime` is a naive local string (no offset); use a Windows time zone name (Graph normalises to `"UTC"` in practice — use `"UTC"` on write, treat the date component as the day-level due date, ignore time-of-day)
  - `completedDateTime`
  - `lastModifiedDateTime` — conflict resolution
  - `isReminderOn` + `reminderDateTime` — set both when creating from Obsidian, using the due date at a fixed local time (default 08:00, configurable) so it actually nudges. (Flag: whether this reliably fires a client notification depends on the To Do/Outlook client being installed and running — verify in testing, not guaranteed by Graph alone.)
- Delta query (`/me/todo/lists/{id}/tasks/delta`) available and should be used for polling instead of full list fetch each time.
- No documented per-service throttling limit; global cap (130k req/10s/app) applies, irrelevant at this scale.

## Auth

- Entra ID app registration: public client (no secret), device code flow (avoids embedded-browser/CORS complexity in Obsidian's renderer sandbox).
- MSAL (browser/node-agnostic device-code grant against `/common` or the specific tenant, whichever the user's app registration targets).
- Scope requested: `Tasks.ReadWrite`.
- Flow: "Sign in to Microsoft" command shows a modal with the device code + verification URL (user completes in their own browser), then polls token endpoint until granted. Access + refresh token cached; access token silently refreshed via MSAL's refresh flow using the cached refresh token.

## Token storage

Use Obsidian's built-in `app.secretStorage` (available since Obsidian 1.11.4; OS-encrypted at rest — DPAPI/Keychain/libsecret depending on platform under the hood).

- `minAppVersion`: `1.11.4`.
- Store the MSAL token cache blob under a prefixed secret id, e.g. `obsidian-todo-sync-msal-cache` (secret ids are a single shared namespace across all plugins — prefix to avoid collisions).
- On `getSecret` returning empty (first run, or pre-1.11.4 vault opened once and then upgraded), prompt re-auth via the "Sign in" command.
- Do not fall back to plaintext `data.json` for the token — if `secretStorage` is unavailable, sync is simply unavailable until the user upgrades Obsidian (surfaced as a settings-tab notice).

## Sync-identity marker

Each synced Obsidian task line gets an invisible HTML comment appended holding the To Do task id:

```
- [ ] Renew passport 📅 2026-10-01 %%todo:AAMkAGI1...%%
```

- Chosen over a visible block ref (`^id`) to avoid cluttering rendered/queried task text, and over an external JSON map keyed by file+line (fragile — breaks on any reordering or edit above the line).
- Parser strips the `%%...%%` comment before interpreting the visible task text; only the plugin reads/writes it.

## Parsing

The Tasks-plugin's public `apiV1` only supports UI-automation (create/edit modal, toggle-done command) — it does not expose reading or querying existing tasks. The plugin hand-rolls a minimal parser for its own subset of syntax:

- Checkbox state: `- [ ]` / `- [x]`.
- Due date: `📅 YYYY-MM-DD`.
- Done date: `✅ YYYY-MM-DD` (written on completion, read to detect already-synced completions).
- Recurrence marker: `🔁 ...` — presence causes the task to be skipped from sync entirely (logged once, not repeated per poll).
- Sync marker: `%%todo:<id>%%` (this plugin's own addition, not Tasks-plugin syntax).
- Everything else (description, priority emoji, tags, wikilinks) is treated as opaque title text, passed through to To Do's `title` field as-is and not otherwise interpreted.

Scanning: on load and on each poll, scan all markdown files for lines matching `- [ ]`/`- [x]` with a `📅` due date. (No vault-wide caching needed at this scale — a full vault scan is cheap; avoid premature optimisation. Revisit only if it proves slow in practice.)

## Sync engine

- Poll interval: 10 minutes (`registerInterval`), plus a manual "Sync now" command.
- Each poll:
  1. Delta-fetch changed To Do tasks since last sync cursor (persisted delta token).
  2. Scan vault for due-date task lines (see Parsing).
  3. Match by `%%todo:<id>%%` marker where present.
  4. For matched pairs: compare `lastModifiedDateTime` (remote) against the file's mtime at last-known-sync (persisted per task in `data.json`) to decide which side is newer; last-write-wins, apply the newer side's title/due-date/status to the other.
  5. For local tasks with a due date and no marker: create a new To Do task, append the marker to the line.
  6. For remote tasks with no matching marker anywhere in the vault: queue as "new from remote" (see below).
  7. Persist updated delta cursor and per-task last-synced state.
- Writing to vault files goes through Obsidian's `Vault.process`/editor API (never raw `fs`), to stay consistent with concurrent user edits and avoid corrupting files mid-edit.

## New tasks from To Do -> Obsidian

- A To Do task with no `%%todo:<id>%%` marker anywhere in the vault is new-from-remote.
- Target: current ISO-week weekly note (`Calendar/Weekly/YYYY-Www.md`), "Other Action Items" section, as a new `- [ ] <title> 📅 <due date> %%todo:<id>%%` line.
- If that file doesn't exist yet (weekly note not created from template for the current week), the task is held in a pending queue in `data.json` and flushed in on the next poll where the file is found to exist. The plugin never creates the weekly note itself (Templater-generated, out of scope to fabricate).
- "Other Action Items" section is located by heading match (`### Other Action Items`); the placeholder `- [ ] ✅` line (if still present/empty) is replaced, otherwise the new line is appended under the heading.

## Settings

- Sign in / sign out (shows account name once signed in).
- To Do list selector (defaults to auto-created "Obsidian" list).
- Poll interval (default 10 min).
- Default reminder time-of-day for new tasks pushed to To Do (default 08:00).

## Commands

- `Sign in to Microsoft`
- `Sign out`
- `Sync now`

## Error handling

- Auth failure / expired refresh token: surface a persistent notice, disable polling until re-auth.
- Network failure during poll: log, skip this cycle, retry next interval (no aggressive retry loop).
- 429/throttling: respect `Retry-After`, back off that poll cycle.
- Conflicting edit where both sides changed since last sync: last-write-wins per the timestamp comparison above; no merge UI in v1.

## Module layout

Per repo conventions (`src/main.ts` minimal lifecycle):

```
src/
  main.ts              # lifecycle, registers commands/interval
  settings.ts          # settings interface, defaults, tab
  auth/
    msal-device-code.ts
    token-storage.ts   # app.secretStorage wrapper
  todo-api/
    client.ts          # Graph To Do REST calls
    types.ts
  obsidian-tasks/
    parser.ts           # line <-> task-line data parsing
    weekly-note.ts       # locate/insert into "Other Action Items"
  sync/
    engine.ts            # poll loop, matching, conflict resolution
    pending-queue.ts      # new-from-remote holding queue
  ui/
    device-code-modal.ts
  types.ts
```

## Open items to verify during implementation (not blocking the spec)

- Whether `reminderDateTime` reliably fires a To Do/Outlook client notification in Frank's tenant.
- Whether the tenant's consent policy allows user consent for `Tasks.ReadWrite` without admin approval.
- Whether Graph enforces `@odata.etag`/`If-Match` on `PATCH` (affects whether an extra conflict-detection layer beyond `lastModifiedDateTime` comparison is worth adding later).
