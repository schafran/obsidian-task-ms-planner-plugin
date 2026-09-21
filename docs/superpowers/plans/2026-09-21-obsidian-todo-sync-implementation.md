# Obsidian <-> Microsoft To Do Sync Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Bidirectional sync between due-date Obsidian tasks (Tasks-plugin syntax) and a dedicated Microsoft To Do list, so due tasks get real OS/Outlook reminders.

**Architecture:** Pure-logic modules (parser, weekly-note insertion, Graph client, MSAL cache plugin, pending queue, sync engine) are dependency-injected and unit-testable without the `obsidian` runtime. Thin adapters (`vault-adapter.ts`, `settings.ts`, `ui/device-code-modal.ts`, `main.ts`) wire those modules to the real Obsidian `App`/`Plugin` API and are covered by lighter shape tests plus manual verification, per project convention that UI/runtime behaviour needs hands-on checking.

**Tech Stack:** TypeScript (strict), esbuild bundler (existing), Vitest for unit tests (new dependency), `@azure/msal-node` for device-code auth (new dependency), global `fetch` for Graph calls (no HTTP client dependency needed — Node 18+ and Electron both provide it).

**Spec:** `docs/superpowers/specs/2026-09-21-obsidian-todo-sync-design.md`

## Global Constraints

- Desktop only: `manifest.json` `isDesktopOnly: true`.
- `minAppVersion`: `1.11.4` (required for `app.secretStorage`).
- Only tasks with a `📅 YYYY-MM-DD` due date are synced; undated and `🔁` recurring tasks are skipped.
- Never write vault files via raw `fs` — only `Vault.process`/editor API.
- Never store the MSAL token cache in plaintext `data.json` — only `app.secretStorage`, under secret id `obsidian-todo-sync-msal-cache`.
- Poll interval default 10 minutes via `registerInterval`; also a manual "Sync now" command.
- Last-write-wins conflict resolution using `lastModifiedDateTime` (remote) vs. file mtime at last sync (local) — no merge UI.
- British English spelling in all UI copy, comments, and commit messages.
- No Claude/Claude Code attribution in commits.

---

## Task 0: Test tooling and plugin identity

**Files:**
- Modify: `package.json`
- Create: `vitest.config.ts`
- Create: `tests/mocks/obsidian.ts`
- Create: `tests/setup.smoke.test.ts`
- Modify: `manifest.json`

**Interfaces:**
- Produces: `tests/mocks/obsidian.ts` exports fake runtime classes (`Notice`, `Modal`, `Plugin`, `PluginSettingTab`, `Setting`, `TFile`, `normalizePath`) used by every later task that needs to unit-test code importing `obsidian` at runtime. Real `obsidian` npm package ships only `.d.ts` (no runtime JS) — importing its runtime exports in a test without this alias throws `undefined is not a constructor`.

- [ ] **Step 1: Add Vitest and the MSAL dependency**

```bash
npm install --save-dev vitest
npm install @azure/msal-node
```

- [ ] **Step 2: Add test script to `package.json`**

Edit `package.json` `scripts`:

```json
"scripts": {
  "dev": "node esbuild.config.mjs",
  "build": "tsc -noEmit -skipLibCheck && node esbuild.config.mjs production",
  "test": "vitest run",
  "test:watch": "vitest",
  "version": "node version-bump.mjs && git add manifest.json versions.json",
  "lint": "eslint ."
}
```

- [ ] **Step 3: Write the obsidian runtime mock**

Create `tests/mocks/obsidian.ts`:

```ts
export class Notice {
	constructor(_message: string) {}
}

export class Component {
	load() {}
	unload() {}
	registerEvent(..._args: unknown[]) {}
	registerDomEvent(..._args: unknown[]) {}
	registerInterval(id: number) {
		return id;
	}
}

export class Plugin extends Component {
	app: unknown;
	manifest: unknown;
	constructor(app: unknown, manifest: unknown) {
		super();
		this.app = app;
		this.manifest = manifest;
	}
	addCommand(..._args: unknown[]) {}
	addSettingTab(..._args: unknown[]) {}
	addRibbonIcon(..._args: unknown[]) {}
	addStatusBarItem() {
		return { setText(_t: string) {} };
	}
	async loadData(): Promise<unknown> {
		return {};
	}
	async saveData(_data: unknown): Promise<void> {}
}

export class Modal {
	app: unknown;
	contentEl: HTMLElement;
	constructor(app: unknown) {
		this.app = app;
		this.contentEl = document.createElement('div');
	}
	open() {
		this.onOpen();
	}
	close() {
		this.onClose();
	}
	onOpen() {}
	onClose() {}
}

export class PluginSettingTab {
	app: unknown;
	plugin: unknown;
	containerEl: HTMLElement;
	constructor(app: unknown, plugin: unknown) {
		this.app = app;
		this.plugin = plugin;
		this.containerEl = document.createElement('div');
	}
	display() {}
}

export class Setting {
	settingEl: HTMLElement;
	constructor(containerEl: HTMLElement) {
		this.settingEl = document.createElement('div');
		containerEl.appendChild(this.settingEl);
	}
	setName(_n: string) {
		return this;
	}
	setDesc(_d: string) {
		return this;
	}
	addText(cb: (t: unknown) => void) {
		cb({ setValue: () => ({}), onChange: () => ({}), setPlaceholder: () => ({}) });
		return this;
	}
	addButton(cb: (b: unknown) => void) {
		cb({ setButtonText: () => ({ onClick: () => ({}) }), onClick: () => ({}) });
		return this;
	}
	addToggle(cb: (t: unknown) => void) {
		cb({ setValue: () => ({}), onChange: () => ({}) });
		return this;
	}
}

export class TFile {
	path: string;
	stat: { mtime: number };
	constructor(path: string, mtime = 0) {
		this.path = path;
		this.stat = { mtime };
	}
}

export function normalizePath(path: string): string {
	return path.replace(/\\/g, '/');
}

export interface SecretStorage {
	setSecret(id: string, secret: string): void;
	getSecret(id: string): string | null;
	listSecrets(): string[];
}
```

- [ ] **Step 4: Point Vitest at the mock**

Create `vitest.config.ts`:

```ts
import { defineConfig } from 'vitest/config';
import path from 'node:path';

export default defineConfig({
	test: {
		environment: 'jsdom',
		alias: {
			obsidian: path.resolve(__dirname, 'tests/mocks/obsidian.ts'),
		},
	},
});
```

- [ ] **Step 5: Add jsdom dependency (Vitest's `environment: 'jsdom'` needs it)**

```bash
npm install --save-dev jsdom
```

- [ ] **Step 6: Write a smoke test and confirm the harness works**

Create `tests/setup.smoke.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { Notice } from 'obsidian';

describe('test harness', () => {
	it('resolves the obsidian mock instead of the real (typings-only) package', () => {
		expect(() => new Notice('hello')).not.toThrow();
	});
});
```

Run: `npm test`
Expected: 1 passed.

- [ ] **Step 7: Update plugin identity**

Edit `manifest.json`:

```json
{
	"id": "obsidian-todo-sync",
	"name": "Microsoft To Do Sync",
	"version": "0.1.0",
	"minAppVersion": "1.11.4",
	"description": "Bidirectional sync between due-date Obsidian tasks and a Microsoft To Do list, with real reminders.",
	"author": "Frank Schaufelberger",
	"isDesktopOnly": true
}
```

Edit `package.json` `name`/`description`/`version` to match (`"name": "obsidian-todo-sync"`, `"version": "0.1.0"`, description matching manifest).

- [ ] **Step 8: Commit**

```bash
git add package.json package-lock.json vitest.config.ts tests/mocks/obsidian.ts tests/setup.smoke.test.ts manifest.json
git commit -m "Add Vitest harness with obsidian runtime mock; set plugin identity"
```

---

## Task 1: Shared types

**Files:**
- Create: `src/types.ts`
- Create: `src/todo-api/types.ts`

**Interfaces:**
- Produces: `ParsedTaskLine` (consumed by Task 2, 10), `TodoTask`/`TodoList`/`NewTaskInput`/`DeltaResult`/`RetryAfterError` (consumed by Task 4, 10).

No test needed — pure type declarations, nothing to assert against. (Type-only files are exempt from the "write the test first" step: `tsc -noEmit` in the build script is the check.)

- [ ] **Step 1: Write `src/types.ts`**

```ts
export interface ParsedTaskLine {
	checked: boolean;
	title: string;
	dueDate: string | null; // YYYY-MM-DD
	doneDate: string | null; // YYYY-MM-DD
	recurring: boolean;
	todoId: string | null;
}
```

- [ ] **Step 2: Write `src/todo-api/types.ts`**

```ts
export interface TodoList {
	id: string;
	displayName: string;
}

export type TodoTaskStatus =
	| 'notStarted'
	| 'inProgress'
	| 'completed'
	| 'waitingOnOthers'
	| 'deferred';

export interface TodoDateTime {
	dateTime: string;
	timeZone: string;
}

export interface TodoTask {
	id: string;
	title: string;
	status: TodoTaskStatus;
	dueDateTime: TodoDateTime | null;
	completedDateTime: TodoDateTime | null;
	lastModifiedDateTime: string;
	isReminderOn: boolean;
	reminderDateTime: TodoDateTime | null;
}

export interface NewTaskInput {
	title: string;
	dueDate: string; // YYYY-MM-DD
	reminderTime: string; // HH:mm, local
}

export interface DeltaResult {
	tasks: TodoTask[];
	deltaLink: string;
}

export class RetryAfterError extends Error {
	constructor(public readonly retryAfterSeconds: number) {
		super(`Graph request throttled, retry after ${retryAfterSeconds}s`);
		this.name = 'RetryAfterError';
	}
}
```

- [ ] **Step 3: Verify it compiles**

Run: `npx tsc -noEmit -skipLibCheck`
Expected: no errors.

- [ ] **Step 4: Commit**

```bash
git add src/types.ts src/todo-api/types.ts
git commit -m "Add shared task and Graph To Do types"
```

---

## Task 2: Task-line parser

**Files:**
- Create: `src/obsidian-tasks/parser.ts`
- Test: `src/obsidian-tasks/parser.test.ts`

**Interfaces:**
- Consumes: `ParsedTaskLine` from `src/types.ts` (Task 1).
- Produces: `parseTaskLine(line: string): ParsedTaskLine | null`, `renderTaskLine(task: ParsedTaskLine): string` — consumed by Task 10 (sync engine).

Design decision (documented here, not deferred): the spec treats "everything else" (tags, priority emoji, wikilinks) as opaque title text. Rather than surgically patching the due-date/done-date/marker tokens in place on re-serialisation (fragile against reordering), `renderTaskLine` reconstructs the line in one canonical order: `- [ ] {title} 📅 {due}{ ✅ done}{ %%todo:id%%}`. Opaque title content is preserved verbatim; only the ordering of the plugin's own metadata suffixes is canonicalised. This is simpler and covers every case the sync engine needs.

- [ ] **Step 1: Write the failing tests**

Create `src/obsidian-tasks/parser.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { parseTaskLine, renderTaskLine } from './parser';

describe('parseTaskLine', () => {
	it('returns null for a non-checkbox line', () => {
		expect(parseTaskLine('Just some text')).toBeNull();
	});

	it('returns null for a checkbox line with no due date', () => {
		expect(parseTaskLine('- [ ] Undated task')).toBeNull();
	});

	it('parses an open task with a due date', () => {
		expect(parseTaskLine('- [ ] Renew passport 📅 2026-10-01')).toEqual({
			checked: false,
			title: 'Renew passport',
			dueDate: '2026-10-01',
			doneDate: null,
			recurring: false,
			todoId: null,
		});
	});

	it('parses a completed task with a done date', () => {
		expect(
			parseTaskLine('- [x] Renew passport 📅 2026-10-01 ✅ 2026-09-30'),
		).toEqual({
			checked: true,
			title: 'Renew passport',
			dueDate: '2026-10-01',
			doneDate: '2026-09-30',
			recurring: false,
			todoId: null,
		});
	});

	it('parses the sync marker', () => {
		const parsed = parseTaskLine(
			'- [ ] Renew passport 📅 2026-10-01 %%todo:AAMkAGI1%%',
		);
		expect(parsed?.todoId).toBe('AAMkAGI1');
		expect(parsed?.title).toBe('Renew passport');
	});

	it('flags recurrence and does not treat 🔁 as part of the title', () => {
		const parsed = parseTaskLine('- [ ] Water plants 📅 2026-10-01 🔁 every week');
		expect(parsed?.recurring).toBe(true);
	});

	it('preserves tags and wikilinks as opaque title text', () => {
		const parsed = parseTaskLine(
			'- [ ] Review [[Design doc]] #urgent 📅 2026-10-01',
		);
		expect(parsed?.title).toBe('Review [[Design doc]] #urgent');
	});

	it('is case-insensitive on the checkbox mark', () => {
		expect(parseTaskLine('- [X] Done 📅 2026-10-01')?.checked).toBe(true);
	});
});

describe('renderTaskLine', () => {
	it('renders an open task with a due date', () => {
		expect(
			renderTaskLine({
				checked: false,
				title: 'Renew passport',
				dueDate: '2026-10-01',
				doneDate: null,
				recurring: false,
				todoId: null,
			}),
		).toBe('- [ ] Renew passport 📅 2026-10-01');
	});

	it('renders a completed task with the marker', () => {
		expect(
			renderTaskLine({
				checked: true,
				title: 'Renew passport',
				dueDate: '2026-10-01',
				doneDate: '2026-09-30',
				recurring: false,
				todoId: 'AAMkAGI1',
			}),
		).toBe('- [x] Renew passport 📅 2026-10-01 ✅ 2026-09-30 %%todo:AAMkAGI1%%');
	});

	it('round-trips parse -> render for a marker-bearing line', () => {
		const line = '- [ ] Review [[Design doc]] #urgent 📅 2026-10-01 %%todo:abc%%';
		const parsed = parseTaskLine(line);
		expect(parsed).not.toBeNull();
		expect(renderTaskLine(parsed!)).toBe(line);
	});
});
```

- [ ] **Step 2: Run tests, confirm they fail**

Run: `npx vitest run src/obsidian-tasks/parser.test.ts`
Expected: FAIL — `Cannot find module './parser'`.

- [ ] **Step 3: Implement the parser**

Create `src/obsidian-tasks/parser.ts`:

```ts
import type { ParsedTaskLine } from '../types';

const CHECKBOX_RE = /^-\s\[([ xX])\]\s(.*)$/;
const DUE_DATE_RE = /📅\s*(\d{4}-\d{2}-\d{2})/;
const DONE_DATE_RE = /✅\s*(\d{4}-\d{2}-\d{2})/;
const RECURRENCE_RE = /🔁[^\n]*$/;
const MARKER_RE = /%%todo:([^%]+)%%/;

export function parseTaskLine(line: string): ParsedTaskLine | null {
	const checkboxMatch = CHECKBOX_RE.exec(line.trim());
	if (!checkboxMatch) return null;

	const [, mark, rest] = checkboxMatch;
	const dueMatch = DUE_DATE_RE.exec(rest);
	if (!dueMatch) return null;

	const doneMatch = DONE_DATE_RE.exec(rest);
	const markerMatch = MARKER_RE.exec(rest);
	const recurring = RECURRENCE_RE.test(rest);

	const title = rest
		.replace(MARKER_RE, '')
		.replace(RECURRENCE_RE, '')
		.replace(DONE_DATE_RE, '')
		.replace(DUE_DATE_RE, '')
		.trim();

	return {
		checked: mark.toLowerCase() === 'x',
		title,
		dueDate: dueMatch[1],
		doneDate: doneMatch ? doneMatch[1] : null,
		recurring,
		todoId: markerMatch ? markerMatch[1] : null,
	};
}

export function renderTaskLine(task: ParsedTaskLine): string {
	const mark = task.checked ? 'x' : ' ';
	let line = `- [${mark}] ${task.title} 📅 ${task.dueDate}`;
	if (task.doneDate) line += ` ✅ ${task.doneDate}`;
	if (task.todoId) line += ` %%todo:${task.todoId}%%`;
	return line;
}
```

- [ ] **Step 4: Run tests, confirm they pass**

Run: `npx vitest run src/obsidian-tasks/parser.test.ts`
Expected: PASS, 10 tests.

- [ ] **Step 5: Commit**

```bash
git add src/obsidian-tasks/parser.ts src/obsidian-tasks/parser.test.ts
git commit -m "Add Tasks-plugin line parser and renderer"
```

---

## Task 3: Weekly note "Other Action Items" insertion

**Files:**
- Create: `src/obsidian-tasks/weekly-note.ts`
- Test: `src/obsidian-tasks/weekly-note.test.ts`

**Interfaces:**
- Produces: `getIsoWeekNotePath(date: Date): string`, `insertActionItem(content: string, newLine: string): string` — consumed by Task 10.

- [ ] **Step 1: Write the failing tests**

Create `src/obsidian-tasks/weekly-note.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { getIsoWeekNotePath, insertActionItem } from './weekly-note';

describe('getIsoWeekNotePath', () => {
	it('computes the ISO week for a Thursday (unambiguous case)', () => {
		// 2026-01-01 is a Thursday -> ISO week 1 of 2026.
		expect(getIsoWeekNotePath(new Date(Date.UTC(2026, 0, 1)))).toBe(
			'Calendar/Weekly/2026-W01.md',
		);
	});

	it('rolls a Sunday into the same ISO week as its Monday', () => {
		// 2026-09-21 is a Monday (ISO week 39); 2026-09-27 is the following Sunday.
		expect(getIsoWeekNotePath(new Date(Date.UTC(2026, 8, 21)))).toBe(
			'Calendar/Weekly/2026-W39.md',
		);
		expect(getIsoWeekNotePath(new Date(Date.UTC(2026, 8, 27)))).toBe(
			'Calendar/Weekly/2026-W39.md',
		);
	});

	it('assigns late-December dates to week 1 of the next year when applicable', () => {
		// 2025-12-31 is a Wednesday in the ISO week containing 2026-01-01.
		expect(getIsoWeekNotePath(new Date(Date.UTC(2025, 11, 31)))).toBe(
			'Calendar/Weekly/2026-W01.md',
		);
	});
});

describe('insertActionItem', () => {
	const NEW_LINE = '- [ ] Buy milk 📅 2026-10-01 %%todo:xyz%%';

	it('replaces the empty placeholder line under the heading', () => {
		const content = [
			'# Week 39',
			'',
			'### Other Action Items',
			'- [ ] ✅',
			'',
			'### Notes',
		].join('\n');
		const result = insertActionItem(content, NEW_LINE);
		expect(result).toBe(
			['# Week 39', '', '### Other Action Items', NEW_LINE, '', '### Notes'].join(
				'\n',
			),
		);
	});

	it('appends after existing items when there is no placeholder', () => {
		const content = [
			'### Other Action Items',
			'- [ ] Existing item 📅 2026-09-25',
			'### Notes',
		].join('\n');
		const result = insertActionItem(content, NEW_LINE);
		expect(result).toBe(
			[
				'### Other Action Items',
				'- [ ] Existing item 📅 2026-09-25',
				NEW_LINE,
				'### Notes',
			].join('\n'),
		);
	});

	it('appends directly under the heading when the section is empty and heading is last line', () => {
		const content = ['# Week 39', '', '### Other Action Items'].join('\n');
		const result = insertActionItem(content, NEW_LINE);
		expect(result).toBe(
			['# Week 39', '', '### Other Action Items', NEW_LINE].join('\n'),
		);
	});

	it('throws if the heading is missing (caller must have checked the file exists and matches)', () => {
		expect(() => insertActionItem('# Week 39\nno such section', NEW_LINE)).toThrow(
			/Other Action Items/,
		);
	});
});
```

- [ ] **Step 2: Run tests, confirm they fail**

Run: `npx vitest run src/obsidian-tasks/weekly-note.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/obsidian-tasks/weekly-note.ts`:

```ts
const HEADING = '### Other Action Items';
const PLACEHOLDER = '- [ ] ✅';

export function getIsoWeekNotePath(date: Date): string {
	const d = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
	const dayNum = (d.getUTCDay() + 6) % 7; // Monday = 0
	d.setUTCDate(d.getUTCDate() - dayNum + 3); // move to this ISO week's Thursday
	const isoYear = d.getUTCFullYear();
	const firstThursday = new Date(Date.UTC(isoYear, 0, 4));
	const firstDayNum = (firstThursday.getUTCDay() + 6) % 7;
	firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDayNum + 3);
	const week =
		1 + Math.round((d.getTime() - firstThursday.getTime()) / (7 * 24 * 3600 * 1000));
	return `Calendar/Weekly/${isoYear}-W${String(week).padStart(2, '0')}.md`;
}

export function insertActionItem(content: string, newLine: string): string {
	const lines = content.split('\n');
	const headingIndex = lines.findIndex((l) => l.trim() === HEADING);
	if (headingIndex === -1) {
		throw new Error(`"${HEADING}" heading not found in note`);
	}

	let sectionEnd = lines.length;
	for (let i = headingIndex + 1; i < lines.length; i++) {
		if (lines[i]!.trim().startsWith('#')) {
			sectionEnd = i;
			break;
		}
	}

	const placeholderIndex = lines
		.slice(headingIndex + 1, sectionEnd)
		.findIndex((l) => l.trim() === PLACEHOLDER);

	if (placeholderIndex !== -1) {
		lines[headingIndex + 1 + placeholderIndex] = newLine;
		return lines.join('\n');
	}

	let lastContentIndex = headingIndex;
	for (let i = headingIndex + 1; i < sectionEnd; i++) {
		if (lines[i]!.trim() !== '') lastContentIndex = i;
	}
	lines.splice(lastContentIndex + 1, 0, newLine);
	return lines.join('\n');
}
```

- [ ] **Step 4: Run tests, confirm they pass**

Run: `npx vitest run src/obsidian-tasks/weekly-note.test.ts`
Expected: PASS, 7 tests.

- [ ] **Step 5: Commit**

```bash
git add src/obsidian-tasks/weekly-note.ts src/obsidian-tasks/weekly-note.test.ts
git commit -m "Add ISO week note path and Other Action Items insertion"
```

---

## Task 4: Graph To Do API client

**Files:**
- Create: `src/todo-api/client.ts`
- Test: `src/todo-api/client.test.ts`

**Interfaces:**
- Consumes: `TodoTask`, `TodoList`, `NewTaskInput`, `DeltaResult`, `RetryAfterError` from `src/todo-api/types.ts` (Task 1).
- Produces: `TodoClient` class with `ensureList`, `createTask`, `updateTask`, `fetchDelta` — consumed by Task 10.

- [ ] **Step 1: Write the failing tests**

Create `src/todo-api/client.test.ts`:

```ts
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { TodoClient } from './client';
import { RetryAfterError } from './types';

function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}) {
	return {
		ok: status >= 200 && status < 300,
		status,
		headers: { get: (k: string) => headers[k] ?? null },
		json: async () => body,
		text: async () => JSON.stringify(body),
	} as Response;
}

describe('TodoClient', () => {
	const getAccessToken = vi.fn(async () => 'fake-token');
	let fetchMock: ReturnType<typeof vi.fn>;

	beforeEach(() => {
		fetchMock = vi.fn();
		vi.stubGlobal('fetch', fetchMock);
	});

	afterEach(() => {
		vi.unstubAllGlobals();
		getAccessToken.mockClear();
	});

	it('ensureList returns an existing list by display name without creating one', async () => {
		fetchMock.mockResolvedValueOnce(
			jsonResponse({ value: [{ id: '1', displayName: 'Obsidian' }] }),
		);
		const client = new TodoClient(getAccessToken);
		const list = await client.ensureList('Obsidian');
		expect(list).toEqual({ id: '1', displayName: 'Obsidian' });
		expect(fetchMock).toHaveBeenCalledTimes(1);
	});

	it('ensureList creates the list when none matches', async () => {
		fetchMock
			.mockResolvedValueOnce(jsonResponse({ value: [] }))
			.mockResolvedValueOnce(jsonResponse({ id: '2', displayName: 'Obsidian' }));
		const client = new TodoClient(getAccessToken);
		const list = await client.ensureList('Obsidian');
		expect(list).toEqual({ id: '2', displayName: 'Obsidian' });
		expect(fetchMock).toHaveBeenCalledTimes(2);
		const [, createInit] = fetchMock.mock.calls[1]!;
		expect(JSON.parse(createInit.body)).toEqual({ displayName: 'Obsidian' });
	});

	it('createTask sends title, due date, and reminder', async () => {
		fetchMock.mockResolvedValueOnce(jsonResponse({ id: 't1' }));
		const client = new TodoClient(getAccessToken);
		await client.createTask('list1', {
			title: 'Renew passport',
			dueDate: '2026-10-01',
			reminderTime: '08:00',
		});
		const [url, init] = fetchMock.mock.calls[0]!;
		expect(url).toBe('https://graph.microsoft.com/v1.0/me/todo/lists/list1/tasks');
		const body = JSON.parse(init.body);
		expect(body.title).toBe('Renew passport');
		expect(body.dueDateTime).toEqual({ dateTime: '2026-10-01T00:00:00', timeZone: 'UTC' });
		expect(body.reminderDateTime).toEqual({
			dateTime: '2026-10-01T08:00:00',
			timeZone: 'UTC',
		});
		expect(body.isReminderOn).toBe(true);
	});

	it('updateTask PATCHes only the given fields', async () => {
		fetchMock.mockResolvedValueOnce(jsonResponse({ id: 't1', status: 'completed' }));
		const client = new TodoClient(getAccessToken);
		await client.updateTask('list1', 't1', { status: 'completed' });
		const [url, init] = fetchMock.mock.calls[0]!;
		expect(url).toBe('https://graph.microsoft.com/v1.0/me/todo/lists/list1/tasks/t1');
		expect(init.method).toBe('PATCH');
		expect(JSON.parse(init.body)).toEqual({ status: 'completed' });
	});

	it('fetchDelta follows nextLink pages and returns the final deltaLink', async () => {
		fetchMock
			.mockResolvedValueOnce(
				jsonResponse({
					value: [{ id: 'a' }],
					'@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/todo/lists/list1/tasks/delta?$skiptoken=1',
				}),
			)
			.mockResolvedValueOnce(
				jsonResponse({ value: [{ id: 'b' }], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/final' }),
			);
		const client = new TodoClient(getAccessToken);
		const result = await client.fetchDelta('list1');
		expect(result.tasks.map((t) => t.id)).toEqual(['a', 'b']);
		expect(result.deltaLink).toBe('https://graph.microsoft.com/v1.0/final');
	});

	it('throws RetryAfterError on 429 with the Retry-After value', async () => {
		fetchMock.mockResolvedValueOnce(jsonResponse({}, 429, { 'Retry-After': '30' }));
		const client = new TodoClient(getAccessToken);
		await expect(client.fetchDelta('list1')).rejects.toBeInstanceOf(RetryAfterError);
		try {
			await client.fetchDelta('list1');
		} catch (e) {
			expect((e as InstanceType<typeof RetryAfterError>).retryAfterSeconds).toBe(30);
		}
	});

	it('throws a descriptive error on other non-2xx responses', async () => {
		fetchMock.mockResolvedValueOnce(jsonResponse({ error: 'nope' }, 500));
		const client = new TodoClient(getAccessToken);
		await expect(client.fetchDelta('list1')).rejects.toThrow(/500/);
	});
});
```

- [ ] **Step 2: Run tests, confirm they fail**

Run: `npx vitest run src/todo-api/client.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/todo-api/client.ts`:

```ts
import { RetryAfterError, type DeltaResult, type NewTaskInput, type TodoList, type TodoTask } from './types';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export class TodoClient {
	constructor(private readonly getAccessToken: () => Promise<string>) {}

	private async request<T>(pathOrUrl: string, init: RequestInit = {}): Promise<T> {
		const token = await this.getAccessToken();
		const url = pathOrUrl.startsWith('http') ? pathOrUrl : `${GRAPH_BASE}${pathOrUrl}`;
		const res = await fetch(url, {
			...init,
			headers: {
				...(init.headers as Record<string, string> | undefined),
				Authorization: `Bearer ${token}`,
				'Content-Type': 'application/json',
			},
		});

		if (res.status === 429) {
			const retryAfter = Number(res.headers.get('Retry-After') ?? '60');
			throw new RetryAfterError(retryAfter);
		}
		if (!res.ok) {
			throw new Error(`Graph request failed: ${res.status} ${await res.text()}`);
		}
		if (res.status === 204) return undefined as T;
		return (await res.json()) as T;
	}

	async listLists(): Promise<TodoList[]> {
		const data = await this.request<{ value: TodoList[] }>('/me/todo/lists');
		return data.value;
	}

	async ensureList(displayName: string): Promise<TodoList> {
		const lists = await this.listLists();
		const existing = lists.find((l) => l.displayName === displayName);
		if (existing) return existing;
		return this.request<TodoList>('/me/todo/lists', {
			method: 'POST',
			body: JSON.stringify({ displayName }),
		});
	}

	async createTask(listId: string, input: NewTaskInput): Promise<TodoTask> {
		return this.request<TodoTask>(`/me/todo/lists/${listId}/tasks`, {
			method: 'POST',
			body: JSON.stringify({
				title: input.title,
				dueDateTime: { dateTime: `${input.dueDate}T00:00:00`, timeZone: 'UTC' },
				isReminderOn: true,
				reminderDateTime: {
					dateTime: `${input.dueDate}T${input.reminderTime}:00`,
					timeZone: 'UTC',
				},
			}),
		});
	}

	async updateTask(
		listId: string,
		taskId: string,
		patch: Partial<Pick<TodoTask, 'title' | 'status' | 'dueDateTime'>>,
	): Promise<TodoTask> {
		return this.request<TodoTask>(`/me/todo/lists/${listId}/tasks/${taskId}`, {
			method: 'PATCH',
			body: JSON.stringify(patch),
		});
	}

	async fetchDelta(listId: string, deltaLink?: string): Promise<DeltaResult> {
		const path = deltaLink ?? `/me/todo/lists/${listId}/tasks/delta`;
		const data = await this.request<{
			value: TodoTask[];
			'@odata.deltaLink'?: string;
			'@odata.nextLink'?: string;
		}>(path);

		if (data['@odata.nextLink']) {
			const next = await this.fetchDelta(listId, data['@odata.nextLink']);
			return { tasks: [...data.value, ...next.tasks], deltaLink: next.deltaLink };
		}
		return { tasks: data.value, deltaLink: data['@odata.deltaLink'] ?? '' };
	}
}
```

- [ ] **Step 4: Run tests, confirm they pass**

Run: `npx vitest run src/todo-api/client.test.ts`
Expected: PASS, 7 tests.

- [ ] **Step 5: Commit**

```bash
git add src/todo-api/client.ts src/todo-api/client.test.ts
git commit -m "Add Microsoft Graph To Do REST client"
```

---

## Task 5: MSAL cache plugin over `app.secretStorage`

**Files:**
- Create: `src/auth/token-storage.ts`
- Test: `src/auth/token-storage.test.ts`

**Interfaces:**
- Consumes: `SecretStorage` type from `obsidian` (mocked in tests per Task 0).
- Produces: `MSAL_CACHE_SECRET_ID`, `createCachePlugin(secretStorage): ICachePlugin`, `hasStoredSession(secretStorage): boolean` — consumed by Task 6.

- [ ] **Step 1: Write the failing tests**

Create `src/auth/token-storage.test.ts`:

```ts
import { describe, expect, it, vi } from 'vitest';
import type { SecretStorage } from 'obsidian';
import { createCachePlugin, hasStoredSession, MSAL_CACHE_SECRET_ID } from './token-storage';

function fakeSecretStorage(initial: Record<string, string> = {}): SecretStorage {
	const store = { ...initial };
	return {
		setSecret: vi.fn((id: string, secret: string) => {
			store[id] = secret;
		}),
		getSecret: vi.fn((id: string) => store[id] ?? null),
		listSecrets: vi.fn(() => Object.keys(store)),
	};
}

describe('hasStoredSession', () => {
	it('is false when no cache secret has been stored yet', () => {
		expect(hasStoredSession(fakeSecretStorage())).toBe(false);
	});

	it('is true once a cache secret exists', () => {
		expect(
			hasStoredSession(fakeSecretStorage({ [MSAL_CACHE_SECRET_ID]: '{}' })),
		).toBe(true);
	});
});

describe('createCachePlugin', () => {
	it('loads the cache from secret storage on beforeCacheAccess', async () => {
		const secretStorage = fakeSecretStorage({ [MSAL_CACHE_SECRET_ID]: '{"cached":true}' });
		const plugin = createCachePlugin(secretStorage);
		const deserialize = vi.fn();
		await plugin.beforeCacheAccess!({
			tokenCache: { deserialize, serialize: vi.fn() },
			cacheHasChanged: false,
		} as never);
		expect(deserialize).toHaveBeenCalledWith('{"cached":true}');
	});

	it('does nothing on beforeCacheAccess when no cache is stored', async () => {
		const secretStorage = fakeSecretStorage();
		const plugin = createCachePlugin(secretStorage);
		const deserialize = vi.fn();
		await plugin.beforeCacheAccess!({
			tokenCache: { deserialize, serialize: vi.fn() },
			cacheHasChanged: false,
		} as never);
		expect(deserialize).not.toHaveBeenCalled();
	});

	it('persists the cache on afterCacheAccess only if it changed', async () => {
		const secretStorage = fakeSecretStorage();
		const plugin = createCachePlugin(secretStorage);
		const serialize = vi.fn(() => '{"cached":true}');
		await plugin.afterCacheAccess!({
			tokenCache: { deserialize: vi.fn(), serialize },
			cacheHasChanged: true,
		} as never);
		expect(secretStorage.setSecret).toHaveBeenCalledWith(
			MSAL_CACHE_SECRET_ID,
			'{"cached":true}',
		);
	});

	it('does not write when the cache has not changed', async () => {
		const secretStorage = fakeSecretStorage();
		const plugin = createCachePlugin(secretStorage);
		await plugin.afterCacheAccess!({
			tokenCache: { deserialize: vi.fn(), serialize: vi.fn() },
			cacheHasChanged: false,
		} as never);
		expect(secretStorage.setSecret).not.toHaveBeenCalled();
	});
});
```

- [ ] **Step 2: Run tests, confirm they fail**

Run: `npx vitest run src/auth/token-storage.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/auth/token-storage.ts`:

```ts
import type { SecretStorage } from 'obsidian';
import type { ICachePlugin } from '@azure/msal-node';

export const MSAL_CACHE_SECRET_ID = 'obsidian-todo-sync-msal-cache';

export function hasStoredSession(secretStorage: SecretStorage): boolean {
	return secretStorage.getSecret(MSAL_CACHE_SECRET_ID) !== null;
}

export function createCachePlugin(secretStorage: SecretStorage): ICachePlugin {
	return {
		beforeCacheAccess: async (context) => {
			const cached = secretStorage.getSecret(MSAL_CACHE_SECRET_ID);
			if (cached) context.tokenCache.deserialize(cached);
		},
		afterCacheAccess: async (context) => {
			if (context.cacheHasChanged) {
				secretStorage.setSecret(MSAL_CACHE_SECRET_ID, context.tokenCache.serialize());
			}
		},
	};
}
```

- [ ] **Step 4: Run tests, confirm they pass**

Run: `npx vitest run src/auth/token-storage.test.ts`
Expected: PASS, 5 tests.

- [ ] **Step 5: Commit**

```bash
git add src/auth/token-storage.ts src/auth/token-storage.test.ts
git commit -m "Add MSAL cache plugin backed by app.secretStorage"
```

---

## Task 6: MSAL device-code auth

**Files:**
- Create: `src/auth/config.ts`
- Create: `src/auth/msal-device-code.ts`
- Test: `src/auth/msal-device-code.test.ts`

**Interfaces:**
- Consumes: `ICachePlugin` from Task 5 (`createCachePlugin`'s return type).
- Produces: `MsalDeviceCodeAuth` class with `signIn(onDeviceCode)`, `getAccessTokenSilent()`, `signOut()` — consumed by Task 13 (`main.ts`) and Task 11 (device-code modal).

Note: `CLIENT_ID`/`AUTHORITY` in `config.ts` are placeholders for Frank's own Entra ID app registration (public client, no secret, `Tasks.ReadWrite` delegated scope) — filling in the real values is an environment-specific one-time setup step tracked in the spec's "Open items to verify during implementation", not something the sync logic depends on for correctness. Every test in this task mocks `@azure/msal-node` directly, so the real value is irrelevant to what's under test.

- [ ] **Step 1: Write `src/auth/config.ts`**

```ts
// Fill in with the Entra ID app registration created for this plugin
// (public client, no secret, delegated Tasks.ReadWrite scope).
export const CLIENT_ID = 'REPLACE_WITH_ENTRA_APP_CLIENT_ID';
export const AUTHORITY = 'https://login.microsoftonline.com/common';
export const SCOPES = ['Tasks.ReadWrite'];
```

- [ ] **Step 2: Write the failing tests**

Create `src/auth/msal-device-code.test.ts`:

```ts
import { beforeEach, describe, expect, it, vi } from 'vitest';

const acquireTokenByDeviceCode = vi.fn();
const acquireTokenSilent = vi.fn();
const getAllAccounts = vi.fn();
const removeAccount = vi.fn();

vi.mock('@azure/msal-node', () => ({
	PublicClientApplication: vi.fn().mockImplementation(() => ({
		acquireTokenByDeviceCode,
		acquireTokenSilent,
		getTokenCache: () => ({ getAllAccounts, removeAccount }),
	})),
}));

import { MsalDeviceCodeAuth } from './msal-device-code';

describe('MsalDeviceCodeAuth', () => {
	beforeEach(() => {
		acquireTokenByDeviceCode.mockReset();
		acquireTokenSilent.mockReset();
		getAllAccounts.mockReset();
		removeAccount.mockReset();
	});

	it('signIn forwards the device code details to the callback and returns the access token', async () => {
		acquireTokenByDeviceCode.mockImplementation(async ({ deviceCodeCallback }) => {
			deviceCodeCallback({
				userCode: 'ABC123',
				verificationUri: 'https://microsoft.com/devicelogin',
				message: 'Go there and enter ABC123',
				expiresIn: 900,
			});
			return { accessToken: 'token-1', account: { homeAccountId: 'acc-1' } };
		});

		const auth = new MsalDeviceCodeAuth({} as never);
		const onDeviceCode = vi.fn();
		const token = await auth.signIn(onDeviceCode);

		expect(onDeviceCode).toHaveBeenCalledWith({
			userCode: 'ABC123',
			verificationUri: 'https://microsoft.com/devicelogin',
			message: 'Go there and enter ABC123',
			expiresIn: 900,
		});
		expect(token).toBe('token-1');
	});

	it('signIn throws when MSAL returns no result', async () => {
		acquireTokenByDeviceCode.mockResolvedValue(null);
		const auth = new MsalDeviceCodeAuth({} as never);
		await expect(auth.signIn(vi.fn())).rejects.toThrow(/no result/);
	});

	it('getAccessTokenSilent returns null when there is no cached account', async () => {
		getAllAccounts.mockResolvedValue([]);
		const auth = new MsalDeviceCodeAuth({} as never);
		expect(await auth.getAccessTokenSilent()).toBeNull();
	});

	it('getAccessTokenSilent uses the cached account to acquire a token silently', async () => {
		const account = { homeAccountId: 'acc-1' };
		getAllAccounts.mockResolvedValue([account]);
		acquireTokenSilent.mockResolvedValue({ accessToken: 'token-2' });
		const auth = new MsalDeviceCodeAuth({} as never);
		const token = await auth.getAccessTokenSilent();
		expect(token).toBe('token-2');
		expect(acquireTokenSilent).toHaveBeenCalledWith(
			expect.objectContaining({ account }),
		);
	});

	it('getAccessTokenSilent returns null if silent acquisition fails', async () => {
		getAllAccounts.mockResolvedValue([{ homeAccountId: 'acc-1' }]);
		acquireTokenSilent.mockRejectedValue(new Error('interaction_required'));
		const auth = new MsalDeviceCodeAuth({} as never);
		expect(await auth.getAccessTokenSilent()).toBeNull();
	});

	it('signOut removes all cached accounts', async () => {
		const accounts = [{ homeAccountId: 'acc-1' }, { homeAccountId: 'acc-2' }];
		getAllAccounts.mockResolvedValue(accounts);
		const auth = new MsalDeviceCodeAuth({} as never);
		await auth.signOut();
		expect(removeAccount).toHaveBeenCalledTimes(2);
		expect(removeAccount).toHaveBeenCalledWith(accounts[0]);
		expect(removeAccount).toHaveBeenCalledWith(accounts[1]);
	});
});
```

- [ ] **Step 3: Run tests, confirm they fail**

Run: `npx vitest run src/auth/msal-device-code.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 4: Implement**

Create `src/auth/msal-device-code.ts`:

```ts
import { PublicClientApplication, type AccountInfo, type ICachePlugin } from '@azure/msal-node';
import { AUTHORITY, CLIENT_ID, SCOPES } from './config';

export interface DeviceCodeInfo {
	userCode: string;
	verificationUri: string;
	message: string;
	expiresIn: number;
}

export class MsalDeviceCodeAuth {
	private readonly pca: PublicClientApplication;
	private account: AccountInfo | null = null;

	constructor(cachePlugin: ICachePlugin) {
		this.pca = new PublicClientApplication({
			auth: { clientId: CLIENT_ID, authority: AUTHORITY },
			cache: { cachePlugin },
		});
	}

	async signIn(onDeviceCode: (info: DeviceCodeInfo) => void): Promise<string> {
		const result = await this.pca.acquireTokenByDeviceCode({
			scopes: SCOPES,
			deviceCodeCallback: (response) => {
				onDeviceCode({
					userCode: response.userCode,
					verificationUri: response.verificationUri,
					message: response.message,
					expiresIn: response.expiresIn,
				});
			},
		});
		if (!result) throw new Error('Device code sign-in returned no result');
		this.account = result.account;
		return result.accessToken;
	}

	async getAccessTokenSilent(): Promise<string | null> {
		const accounts = await this.pca.getTokenCache().getAllAccounts();
		const account = this.account ?? accounts[0];
		if (!account) return null;
		this.account = account;
		try {
			const result = await this.pca.acquireTokenSilent({ scopes: SCOPES, account });
			return result?.accessToken ?? null;
		} catch {
			return null;
		}
	}

	async signOut(): Promise<void> {
		const cache = this.pca.getTokenCache();
		const accounts = await cache.getAllAccounts();
		for (const account of accounts) {
			await cache.removeAccount(account);
		}
		this.account = null;
	}
}
```

- [ ] **Step 5: Run tests, confirm they pass**

Run: `npx vitest run src/auth/msal-device-code.test.ts`
Expected: PASS, 6 tests.

- [ ] **Step 6: Commit**

```bash
git add src/auth/config.ts src/auth/msal-device-code.ts src/auth/msal-device-code.test.ts
git commit -m "Add MSAL device-code sign-in flow"
```

---

## Task 7: Pending queue (new-from-remote holding area)

**Files:**
- Create: `src/sync/pending-queue.ts`
- Test: `src/sync/pending-queue.test.ts`

**Interfaces:**
- Produces: `PendingRemoteTask` interface, `PendingQueue` class (`add`, `all`, `remove`, `toJSON`, `fromJSON`) — consumed by Task 8 and Task 10.

- [ ] **Step 1: Write the failing tests**

Create `src/sync/pending-queue.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { PendingQueue } from './pending-queue';

describe('PendingQueue', () => {
	it('starts empty', () => {
		expect(new PendingQueue().all).toEqual([]);
	});

	it('adds items and exposes them via all', () => {
		const queue = new PendingQueue();
		queue.add({ todoId: 't1', title: 'Buy milk', dueDate: '2026-10-01' });
		expect(queue.all).toEqual([{ todoId: 't1', title: 'Buy milk', dueDate: '2026-10-01' }]);
	});

	it('removes an item by todoId', () => {
		const queue = new PendingQueue();
		queue.add({ todoId: 't1', title: 'Buy milk', dueDate: '2026-10-01' });
		queue.add({ todoId: 't2', title: 'Water plants', dueDate: '2026-10-02' });
		queue.remove('t1');
		expect(queue.all.map((t) => t.todoId)).toEqual(['t2']);
	});

	it('round-trips through toJSON/fromJSON', () => {
		const queue = new PendingQueue();
		queue.add({ todoId: 't1', title: 'Buy milk', dueDate: '2026-10-01' });
		const restored = PendingQueue.fromJSON(queue.toJSON());
		expect(restored.all).toEqual(queue.all);
	});

	it('fromJSON tolerates undefined (first run, no persisted data yet)', () => {
		expect(PendingQueue.fromJSON(undefined).all).toEqual([]);
	});
});
```

- [ ] **Step 2: Run tests, confirm they fail**

Run: `npx vitest run src/sync/pending-queue.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/sync/pending-queue.ts`:

```ts
export interface PendingRemoteTask {
	todoId: string;
	title: string;
	dueDate: string;
}

export class PendingQueue {
	private items: PendingRemoteTask[];

	constructor(items: PendingRemoteTask[] = []) {
		this.items = items;
	}

	static fromJSON(items: PendingRemoteTask[] | undefined): PendingQueue {
		return new PendingQueue(items ? [...items] : []);
	}

	toJSON(): PendingRemoteTask[] {
		return [...this.items];
	}

	get all(): PendingRemoteTask[] {
		return [...this.items];
	}

	add(task: PendingRemoteTask): void {
		this.items.push(task);
	}

	remove(todoId: string): void {
		this.items = this.items.filter((t) => t.todoId !== todoId);
	}
}
```

- [ ] **Step 4: Run tests, confirm they pass**

Run: `npx vitest run src/sync/pending-queue.test.ts`
Expected: PASS, 5 tests.

- [ ] **Step 5: Commit**

```bash
git add src/sync/pending-queue.ts src/sync/pending-queue.test.ts
git commit -m "Add pending queue for new-from-remote tasks"
```

---

## Task 8: Persisted sync state store

**Files:**
- Create: `src/sync/state-store.ts`
- Test: `src/sync/state-store.test.ts`

**Interfaces:**
- Consumes: `PendingRemoteTask` from Task 7.
- Produces: `TaskSyncState`, `SyncData` interfaces, `loadSyncData(raw)`, `serializeSyncData(data)` — consumed by Task 10 and Task 13 (`main.ts` reads/writes this via `loadData()`/`saveData()`).

- [ ] **Step 1: Write the failing tests**

Create `src/sync/state-store.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { loadSyncData, serializeSyncData } from './state-store';

describe('loadSyncData', () => {
	it('returns empty defaults when given undefined (first run)', () => {
		expect(loadSyncData(undefined)).toEqual({
			deltaLink: '',
			taskStates: {},
			pending: [],
		});
	});

	it('round-trips a populated state through serialize/load', () => {
		const data = {
			deltaLink: 'https://graph.microsoft.com/v1.0/delta-cursor',
			taskStates: {
				t1: { lastKnownRemoteModified: '2026-09-20T10:00:00Z', lastSyncedAtMs: 1758000000000 },
			},
			pending: [{ todoId: 't2', title: 'Buy milk', dueDate: '2026-10-01' }],
		};
		expect(loadSyncData(serializeSyncData(data))).toEqual(data);
	});

	it('drops malformed taskStates entries rather than throwing', () => {
		const raw = {
			deltaLink: '',
			taskStates: { t1: { lastKnownRemoteModified: 123 } },
			pending: [],
		};
		expect(loadSyncData(raw)).toEqual({ deltaLink: '', taskStates: {}, pending: [] });
	});
});
```

- [ ] **Step 2: Run tests, confirm they fail**

Run: `npx vitest run src/sync/state-store.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/sync/state-store.ts`:

```ts
import type { PendingRemoteTask } from './pending-queue';

export interface TaskSyncState {
	lastKnownRemoteModified: string;
	lastSyncedAtMs: number;
}

export interface SyncData {
	deltaLink: string;
	taskStates: Record<string, TaskSyncState>;
	pending: PendingRemoteTask[];
}

const EMPTY: SyncData = { deltaLink: '', taskStates: {}, pending: [] };

function isValidTaskState(value: unknown): value is TaskSyncState {
	if (typeof value !== 'object' || value === null) return false;
	const v = value as Record<string, unknown>;
	return typeof v.lastKnownRemoteModified === 'string' && typeof v.lastSyncedAtMs === 'number';
}

export function loadSyncData(raw: unknown): SyncData {
	if (typeof raw !== 'object' || raw === null) return { ...EMPTY };
	const r = raw as Record<string, unknown>;

	const taskStates: Record<string, TaskSyncState> = {};
	if (typeof r.taskStates === 'object' && r.taskStates !== null) {
		for (const [id, state] of Object.entries(r.taskStates as Record<string, unknown>)) {
			if (isValidTaskState(state)) taskStates[id] = state;
		}
	}

	return {
		deltaLink: typeof r.deltaLink === 'string' ? r.deltaLink : '',
		taskStates,
		pending: Array.isArray(r.pending) ? (r.pending as PendingRemoteTask[]) : [],
	};
}

export function serializeSyncData(data: SyncData): SyncData {
	return {
		deltaLink: data.deltaLink,
		taskStates: { ...data.taskStates },
		pending: [...data.pending],
	};
}
```

- [ ] **Step 4: Run tests, confirm they pass**

Run: `npx vitest run src/sync/state-store.test.ts`
Expected: PASS, 3 tests.

- [ ] **Step 5: Commit**

```bash
git add src/sync/state-store.ts src/sync/state-store.test.ts
git commit -m "Add persisted sync state store with defensive loading"
```

---

## Task 9: Vault adapter

**Files:**
- Create: `src/obsidian-tasks/vault-adapter.ts`

**Interfaces:**
- Produces: `VaultFile { path: string; mtimeMs: number }`, `VaultAdapter` interface (`listMarkdownFiles`, `read`, `update`, `getFile`), `createVaultAdapter(app: App): VaultAdapter` — consumed by Task 10 (as the interface tests fake) and Task 13 (`main.ts` wires the real implementation).

This is a thin wrapper with no branching logic of its own (it delegates straight to `Vault.getMarkdownFiles`/`Vault.read`/`Vault.process`), so per Task Right-Sizing it is not independently unit-tested — Task 10's engine tests exercise the interface via a fake, and Task 14's manual verification exercises the real implementation end-to-end.

- [ ] **Step 1: Implement**

Create `src/obsidian-tasks/vault-adapter.ts`:

```ts
import type { App, TFile } from 'obsidian';

export interface VaultFile {
	path: string;
	mtimeMs: number;
}

export interface VaultAdapter {
	listMarkdownFiles(): VaultFile[];
	read(file: VaultFile): Promise<string>;
	update(file: VaultFile, mutate: (content: string) => string): Promise<void>;
	getFile(path: string): VaultFile | null;
}

function toVaultFile(file: TFile): VaultFile {
	return { path: file.path, mtimeMs: file.stat.mtime };
}

export function createVaultAdapter(app: App): VaultAdapter {
	return {
		listMarkdownFiles() {
			return app.vault.getMarkdownFiles().map(toVaultFile);
		},
		async read(file) {
			const tFile = app.vault.getFileByPath(file.path);
			if (!tFile) throw new Error(`File not found: ${file.path}`);
			return app.vault.read(tFile);
		},
		async update(file, mutate) {
			const tFile = app.vault.getFileByPath(file.path);
			if (!tFile) throw new Error(`File not found: ${file.path}`);
			await app.vault.process(tFile, mutate);
		},
		getFile(path) {
			const tFile = app.vault.getFileByPath(path);
			return tFile ? toVaultFile(tFile) : null;
		},
	};
}
```

- [ ] **Step 2: Verify it compiles**

Run: `npx tsc -noEmit -skipLibCheck`
Expected: no errors. (`getFileByPath` requires Obsidian API >= 1.0; confirm it's present in `node_modules/obsidian/obsidian.d.ts` — if not on this API version, fall back to `app.vault.getAbstractFileByPath(path)` cast to `TFile`.)

- [ ] **Step 3: Commit**

```bash
git add src/obsidian-tasks/vault-adapter.ts
git commit -m "Add thin Vault adapter for the sync engine"
```

---

## Task 10: Sync engine

**Files:**
- Create: `src/sync/engine.ts`
- Test: `src/sync/engine.test.ts`

**Interfaces:**
- Consumes:
  - `parseTaskLine`, `renderTaskLine` (Task 2)
  - `getIsoWeekNotePath`, `insertActionItem` (Task 3)
  - `TodoClient` (Task 4, used via a narrowed `Pick`)
  - `VaultAdapter`, `VaultFile` (Task 9)
  - `SyncData`, `TaskSyncState` (Task 8)
  - `PendingQueue` (Task 7)
- Produces: `SyncEngine` class with `runPollCycle(data: SyncData): Promise<SyncData>` — consumed by Task 13 (`main.ts`'s interval and "Sync now" command).

- [ ] **Step 1: Write the failing tests**

Create `src/sync/engine.test.ts`:

```ts
import { describe, expect, it, vi } from 'vitest';
import { SyncEngine, type SyncEngineDeps } from './engine';
import type { VaultAdapter, VaultFile } from '../obsidian-tasks/vault-adapter';
import type { TodoTask } from '../todo-api/types';
import { loadSyncData } from './state-store';

function fakeVault(files: Record<string, string>): VaultAdapter {
	const store = { ...files };
	const mtimes: Record<string, number> = Object.fromEntries(
		Object.keys(files).map((p) => [p, 1_000]),
	);
	return {
		listMarkdownFiles(): VaultFile[] {
			return Object.keys(store).map((path) => ({ path, mtimeMs: mtimes[path]! }));
		},
		async read(file) {
			return store[file.path]!;
		},
		async update(file, mutate) {
			store[file.path] = mutate(store[file.path]!);
			mtimes[file.path] = Date.now();
		},
		getFile(path) {
			return path in store ? { path, mtimeMs: mtimes[path]! } : null;
		},
		// test-only accessor
		_dump: () => ({ ...store }),
	} as VaultAdapter & { _dump: () => Record<string, string> };
}

function baseTask(overrides: Partial<TodoTask> = {}): TodoTask {
	return {
		id: 't1',
		title: 'Renew passport',
		status: 'notStarted',
		dueDateTime: { dateTime: '2026-10-01T00:00:00', timeZone: 'UTC' },
		completedDateTime: null,
		lastModifiedDateTime: '2026-09-20T09:00:00Z',
		isReminderOn: true,
		reminderDateTime: { dateTime: '2026-10-01T08:00:00', timeZone: 'UTC' },
		...overrides,
	};
}

function deps(overrides: Partial<SyncEngineDeps> = {}): SyncEngineDeps {
	return {
		todo: {
			ensureList: vi.fn(async () => ({ id: 'list1', displayName: 'Obsidian' })),
			createTask: vi.fn(async () => baseTask()),
			updateTask: vi.fn(async () => baseTask()),
			fetchDelta: vi.fn(async () => ({ tasks: [], deltaLink: 'cursor-1' })),
		},
		vault: fakeVault({}),
		now: () => new Date('2026-09-21T12:00:00Z'),
		reminderTime: '08:00',
		listName: 'Obsidian',
		log: vi.fn(),
		...overrides,
	};
}

describe('SyncEngine.runPollCycle', () => {
	it('pushes a new local due-date task to To Do and writes the marker back', async () => {
		const vault = fakeVault({
			'note.md': '- [ ] Renew passport 📅 2026-10-01',
		});
		const createTask = vi.fn(async () => baseTask({ id: 'new-id' }));
		const engine = new SyncEngine(deps({ vault, todo: { ...deps().todo, createTask } }));

		const result = await engine.runPollCycle(loadSyncData(undefined));

		expect(createTask).toHaveBeenCalledWith('list1', {
			title: 'Renew passport',
			dueDate: '2026-10-01',
			reminderTime: '08:00',
		});
		expect((vault as VaultAdapter & { _dump(): Record<string, string> })._dump()['note.md']).toBe(
			'- [ ] Renew passport 📅 2026-10-01 %%todo:new-id%%',
		);
		expect(result.taskStates['new-id']).toBeDefined();
	});

	it('skips recurring tasks entirely', async () => {
		const vault = fakeVault({
			'note.md': '- [ ] Water plants 📅 2026-10-01 🔁 every week',
		});
		const createTask = vi.fn();
		const engine = new SyncEngine(deps({ vault, todo: { ...deps().todo, createTask } }));

		await engine.runPollCycle(loadSyncData(undefined));

		expect(createTask).not.toHaveBeenCalled();
	});

	it('applies a remote completion to the matching local line', async () => {
		const vault = fakeVault({
			'note.md': '- [ ] Renew passport 📅 2026-10-01 %%todo:t1%%',
		});
		const fetchDelta = vi.fn(async () => ({
			tasks: [
				baseTask({
					status: 'completed',
					completedDateTime: { dateTime: '2026-09-21T00:00:00', timeZone: 'UTC' },
					lastModifiedDateTime: '2026-09-21T10:00:00Z',
				}),
			],
			deltaLink: 'cursor-2',
		}));
		const priorData = loadSyncData({
			deltaLink: 'cursor-1',
			taskStates: { t1: { lastKnownRemoteModified: '2026-09-20T09:00:00Z', lastSyncedAtMs: 1_000 } },
			pending: [],
		});
		const engine = new SyncEngine(deps({ vault, todo: { ...deps().todo, fetchDelta } }));

		await engine.runPollCycle(priorData);

		expect((vault as VaultAdapter & { _dump(): Record<string, string> })._dump()['note.md']).toBe(
			'- [x] Renew passport 📅 2026-10-01 ✅ 2026-09-21 %%todo:t1%%',
		);
	});

	it('pushes a local edit to a matched task up to To Do when the local file changed more recently', async () => {
		const vault = fakeVault({
			'note.md': '- [x] Renew passport 📅 2026-10-05 ✅ 2026-09-21 %%todo:t1%%',
		});
		// local file's mtime (from fakeVault) is 1_000; last sync was also at 1_000 in priorData
		// but fakeVault always reports mtimeMs 1_000 for pre-seeded files and bumps on update,
		// so mark this file as changed since last sync by giving it a newer lastSyncedAtMs baseline below 1_000.
		const updateTask = vi.fn(async () => baseTask());
		const priorData = loadSyncData({
			deltaLink: 'cursor-1',
			taskStates: { t1: { lastKnownRemoteModified: '2026-09-20T09:00:00Z', lastSyncedAtMs: 500 } },
			pending: [],
		});
		const engine = new SyncEngine(deps({ vault, todo: { ...deps().todo, updateTask } }));

		await engine.runPollCycle(priorData);

		expect(updateTask).toHaveBeenCalledWith(
			'list1',
			't1',
			expect.objectContaining({
				title: 'Renew passport',
				status: 'completed',
				dueDateTime: { dateTime: '2026-10-05T00:00:00', timeZone: 'UTC' },
			}),
		);
	});

	it('queues an unmatched remote task as pending when the weekly note does not exist yet', async () => {
		const vault = fakeVault({});
		const fetchDelta = vi.fn(async () => ({
			tasks: [baseTask({ id: 'remote-1', title: 'Buy milk' })],
			deltaLink: 'cursor-2',
		}));
		const engine = new SyncEngine(deps({ vault, todo: { ...deps().todo, fetchDelta } }));

		const result = await engine.runPollCycle(loadSyncData(undefined));

		expect(result.pending).toEqual([
			{ todoId: 'remote-1', title: 'Buy milk', dueDate: '2026-10-01' },
		]);
	});

	it('flushes pending tasks into the weekly note once it exists', async () => {
		const vault = fakeVault({
			'Calendar/Weekly/2026-W39.md': ['### Other Action Items', '- [ ] ✅'].join('\n'),
		});
		const engine = new SyncEngine(deps({ vault }));
		const priorData = loadSyncData({
			deltaLink: 'cursor-1',
			taskStates: {},
			pending: [{ todoId: 'remote-1', title: 'Buy milk', dueDate: '2026-10-01' }],
		});

		const result = await engine.runPollCycle(priorData);

		expect(
			(vault as VaultAdapter & { _dump(): Record<string, string> })._dump()[
				'Calendar/Weekly/2026-W39.md'
			],
		).toBe(
			['### Other Action Items', '- [ ] Buy milk 📅 2026-10-01 %%todo:remote-1%%'].join('\n'),
		);
		expect(result.pending).toEqual([]);
	});

	it('persists the new delta cursor from the fetch', async () => {
		const engine = new SyncEngine(deps());
		const result = await engine.runPollCycle(loadSyncData(undefined));
		expect(result.deltaLink).toBe('cursor-1');
	});
});
```

- [ ] **Step 2: Run tests, confirm they fail**

Run: `npx vitest run src/sync/engine.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement**

Create `src/sync/engine.ts`:

```ts
import { parseTaskLine, renderTaskLine } from '../obsidian-tasks/parser';
import { getIsoWeekNotePath, insertActionItem } from '../obsidian-tasks/weekly-note';
import type { VaultAdapter, VaultFile } from '../obsidian-tasks/vault-adapter';
import type { TodoClient } from '../todo-api/client';
import type { TodoTask } from '../todo-api/types';
import type { ParsedTaskLine } from '../types';
import { PendingQueue } from './pending-queue';
import type { SyncData } from './state-store';

export interface SyncEngineDeps {
	todo: Pick<TodoClient, 'ensureList' | 'createTask' | 'updateTask' | 'fetchDelta'>;
	vault: VaultAdapter;
	now: () => Date;
	reminderTime: string;
	listName: string;
	log: (message: string) => void;
}

interface LocatedLine {
	file: VaultFile;
	lineIndex: number;
	parsed: ParsedTaskLine;
}

function toDueDate(task: TodoTask): string {
	return (task.dueDateTime?.dateTime ?? '').slice(0, 10);
}

function toDoneDate(task: TodoTask): string | null {
	return task.completedDateTime ? task.completedDateTime.dateTime.slice(0, 10) : null;
}

export class SyncEngine {
	constructor(private readonly deps: SyncEngineDeps) {}

	async runPollCycle(data: SyncData): Promise<SyncData> {
		const { todo, vault, now, log } = this.deps;
		const list = await todo.ensureList(this.deps.listName);
		const delta = await todo.fetchDelta(list.id, data.deltaLink || undefined);
		const remoteById = new Map(delta.tasks.map((t) => [t.id, t]));

		const files = vault.listMarkdownFiles();
		const localByTodoId = new Map<string, LocatedLine>();
		const localWithoutMarker: LocatedLine[] = [];

		for (const file of files) {
			const content = await vault.read(file);
			const lines = content.split('\n');
			lines.forEach((line, lineIndex) => {
				const parsed = parseTaskLine(line);
				if (!parsed || parsed.recurring) return;
				if (parsed.todoId) {
					localByTodoId.set(parsed.todoId, { file, lineIndex, parsed });
				} else {
					localWithoutMarker.push({ file, lineIndex, parsed });
				}
			});
		}

		const taskStates = { ...data.taskStates };
		const pendingQueue = PendingQueue.fromJSON(data.pending);

		// Matched pairs: decide winner by last-write-wins.
		for (const [todoId, located] of localByTodoId) {
			const remote = remoteById.get(todoId);
			const state = taskStates[todoId];
			if (!remote || !state) continue;

			const remoteChanged = new Date(remote.lastModifiedDateTime).getTime() > state.lastSyncedAtMs;
			const localChanged = located.file.mtimeMs > state.lastSyncedAtMs;

			if (remoteChanged && !localChanged) {
				await this.applyRemoteToLocal(located, remote);
			} else if (localChanged && !remoteChanged) {
				await this.applyLocalToRemote(list.id, located, todo);
			} else if (remoteChanged && localChanged) {
				if (new Date(remote.lastModifiedDateTime).getTime() >= located.file.mtimeMs) {
					await this.applyRemoteToLocal(located, remote);
				} else {
					await this.applyLocalToRemote(list.id, located, todo);
				}
			}

			taskStates[todoId] = {
				lastKnownRemoteModified: remote.lastModifiedDateTime,
				lastSyncedAtMs: now().getTime(),
			};
		}

		// Local tasks with a due date and no marker: push to To Do.
		for (const located of localWithoutMarker) {
			if (!located.parsed.dueDate) continue;
			const created = await todo.createTask(list.id, {
				title: located.parsed.title,
				dueDate: located.parsed.dueDate,
				reminderTime: this.deps.reminderTime,
			});
			await vault.update(located.file, (content) => {
				const lines = content.split('\n');
				lines[located.lineIndex] = renderTaskLine({ ...located.parsed, todoId: created.id });
				return lines.join('\n');
			});
			taskStates[created.id] = {
				lastKnownRemoteModified: created.lastModifiedDateTime,
				lastSyncedAtMs: now().getTime(),
			};
		}

		// Remote tasks with no matching local line: queue as new-from-remote.
		for (const [todoId, remote] of remoteById) {
			if (localByTodoId.has(todoId) || taskStates[todoId]) continue;
			pendingQueue.add({ todoId, title: remote.title, dueDate: toDueDate(remote) });
		}

		// Flush the pending queue into the current weekly note, if it exists.
		const weeklyNotePath = getIsoWeekNotePath(now());
		const weeklyNoteFile = vault.getFile(weeklyNotePath);
		if (weeklyNoteFile) {
			for (const pendingTask of pendingQueue.all) {
				await vault.update(weeklyNoteFile, (content) =>
					insertActionItem(
						content,
						renderTaskLine({
							checked: false,
							title: pendingTask.title,
							dueDate: pendingTask.dueDate,
							doneDate: null,
							recurring: false,
							todoId: pendingTask.todoId,
						}),
					),
				);
				pendingQueue.remove(pendingTask.todoId);
				taskStates[pendingTask.todoId] = {
					lastKnownRemoteModified: now().toISOString(),
					lastSyncedAtMs: now().getTime(),
				};
			}
		}

		log(
			`Sync cycle complete: ${localByTodoId.size} matched, ${localWithoutMarker.length} pushed, ${pendingQueue.all.length} pending.`,
		);

		return { deltaLink: delta.deltaLink, taskStates, pending: pendingQueue.toJSON() };
	}

	private async applyRemoteToLocal(located: LocatedLine, remote: TodoTask): Promise<void> {
		await this.deps.vault.update(located.file, (content) => {
			const lines = content.split('\n');
			lines[located.lineIndex] = renderTaskLine({
				...located.parsed,
				title: remote.title,
				dueDate: toDueDate(remote),
				checked: remote.status === 'completed',
				doneDate: toDoneDate(remote),
			});
			return lines.join('\n');
		});
	}

	private async applyLocalToRemote(
		listId: string,
		located: LocatedLine,
		todo: SyncEngineDeps['todo'],
	): Promise<void> {
		if (!located.parsed.todoId || !located.parsed.dueDate) return;
		await todo.updateTask(listId, located.parsed.todoId, {
			title: located.parsed.title,
			status: located.parsed.checked ? 'completed' : 'notStarted',
			dueDateTime: { dateTime: `${located.parsed.dueDate}T00:00:00`, timeZone: 'UTC' },
		});
	}
}
```

- [ ] **Step 4: Run tests, confirm they pass**

Run: `npx vitest run src/sync/engine.test.ts`
Expected: PASS, 7 tests.

- [ ] **Step 5: Run the full test suite so far**

Run: `npm test`
Expected: all suites pass (parser, weekly-note, client, token-storage, msal-device-code, pending-queue, state-store, engine, smoke).

- [ ] **Step 6: Commit**

```bash
git add src/sync/engine.ts src/sync/engine.test.ts
git commit -m "Add sync engine: matching, last-write-wins, push/pull, pending flush"
```

---

## Task 11: Device-code sign-in modal

**Files:**
- Create: `src/ui/device-code-modal.ts`

**Interfaces:**
- Consumes: `DeviceCodeInfo` from Task 6.
- Produces: `DeviceCodeModal` class — consumed by Task 13 (`main.ts`'s "Sign in to Microsoft" command).

Pure UI glue over the mocked `Modal` base class — no branching logic to unit test beyond "does it render the fields", which is better checked by hand per project convention for UI. Covered by Task 14's manual verification.

- [ ] **Step 1: Implement**

Create `src/ui/device-code-modal.ts`:

```ts
import { App, Modal } from 'obsidian';
import type { DeviceCodeInfo } from '../auth/msal-device-code';

export class DeviceCodeModal extends Modal {
	constructor(app: App) {
		super(app);
	}

	showDeviceCode(info: DeviceCodeInfo): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'Sign in to Microsoft' });
		contentEl.createEl('p', {
			text: `Go to ${info.verificationUri} and enter this code:`,
		});
		contentEl.createEl('p', { text: info.userCode, cls: 'todo-sync-device-code' });
		contentEl.createEl('p', {
			text: 'Waiting for you to complete sign-in in your browser...',
		});
	}

	showSuccess(accountLabel: string): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('p', { text: `Signed in as ${accountLabel}.` });
	}

	showError(message: string): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('p', { text: `Sign-in failed: ${message}` });
	}

	onClose(): void {
		this.contentEl.empty();
	}
}
```

- [ ] **Step 2: Verify it compiles**

Run: `npx tsc -noEmit -skipLibCheck`
Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add src/ui/device-code-modal.ts
git commit -m "Add device-code sign-in modal"
```

---

## Task 12: Settings

**Files:**
- Modify: `src/settings.ts` (replace sample-plugin contents entirely)
- Test: `src/settings.test.ts`

**Interfaces:**
- Produces: `TodoSyncSettings` interface, `DEFAULT_SETTINGS`, `TodoSyncSettingTab` class — consumed by Task 13 (`main.ts`).

- [ ] **Step 1: Write the failing test (defaults only — the tab's rendering is UI glue, covered manually in Task 14)**

Create `src/settings.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { DEFAULT_SETTINGS } from './settings';

describe('DEFAULT_SETTINGS', () => {
	it('defaults to a 10 minute poll interval and 08:00 reminder time', () => {
		expect(DEFAULT_SETTINGS.pollIntervalMinutes).toBe(10);
		expect(DEFAULT_SETTINGS.reminderTime).toBe('08:00');
		expect(DEFAULT_SETTINGS.listName).toBe('Obsidian');
	});
});
```

- [ ] **Step 2: Run test, confirm it fails**

Run: `npx vitest run src/settings.test.ts`
Expected: FAIL — `DEFAULT_SETTINGS` export not found (file still has the old sample-plugin shape).

- [ ] **Step 3: Replace `src/settings.ts`**

```ts
import { App, PluginSettingTab, Setting } from 'obsidian';
import type TodoSyncPlugin from './main';

export interface TodoSyncSettings {
	listName: string;
	pollIntervalMinutes: number;
	reminderTime: string; // HH:mm, local
	signedInAccountLabel: string | null;
}

export const DEFAULT_SETTINGS: TodoSyncSettings = {
	listName: 'Obsidian',
	pollIntervalMinutes: 10,
	reminderTime: '08:00',
	signedInAccountLabel: null,
};

export class TodoSyncSettingTab extends PluginSettingTab {
	constructor(
		app: App,
		private readonly plugin: TodoSyncPlugin,
	) {
		super(app, plugin);
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();

		new Setting(containerEl)
			.setName('Account')
			.setDesc(
				this.plugin.settings.signedInAccountLabel
					? `Signed in as ${this.plugin.settings.signedInAccountLabel}`
					: 'Not signed in',
			)
			.addButton((button) => {
				button
					.setButtonText(
						this.plugin.settings.signedInAccountLabel ? 'Sign out' : 'Sign in',
					)
					.onClick(() => this.plugin.handleSignInOutFromSettings());
			});

		new Setting(containerEl)
			.setName('To Do list name')
			.setDesc('Created automatically on first sign-in if it does not already exist.')
			.addText((text) =>
				text.setValue(this.plugin.settings.listName).onChange(async (value) => {
					this.plugin.settings.listName = value.trim() || DEFAULT_SETTINGS.listName;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl)
			.setName('Poll interval (minutes)')
			.addText((text) =>
				text
					.setValue(String(this.plugin.settings.pollIntervalMinutes))
					.onChange(async (value) => {
						const minutes = Number(value);
						if (Number.isFinite(minutes) && minutes > 0) {
							this.plugin.settings.pollIntervalMinutes = minutes;
							await this.plugin.saveSettings();
							this.plugin.restartPolling();
						}
					}),
			);

		new Setting(containerEl)
			.setName('Default reminder time')
			.setDesc('Local time used for the reminder on tasks pushed from Obsidian to To Do.')
			.addText((text) =>
				text.setValue(this.plugin.settings.reminderTime).onChange(async (value) => {
					if (/^\d{2}:\d{2}$/.test(value)) {
						this.plugin.settings.reminderTime = value;
						await this.plugin.saveSettings();
					}
				}),
			);
	}
}
```

- [ ] **Step 4: Run test, confirm it passes**

Run: `npx vitest run src/settings.test.ts`
Expected: PASS, 1 test.

Note: `src/main.ts` doesn't exist in its final form yet (Task 13 writes it), so this file currently references a not-yet-defined `handleSignInOutFromSettings`/`restartPolling` on `TodoSyncPlugin` — that's fine, it's a type-only import (`import type TodoSyncPlugin from './main'`) and Vitest doesn't type-check. `npx tsc -noEmit` will fail until Task 13 is done; that's expected and resolved by the end of Task 13, not before.

- [ ] **Step 5: Commit**

```bash
git add src/settings.ts src/settings.test.ts
git commit -m "Replace sample settings with To Do sync settings"
```

---

## Task 13: Plugin lifecycle (`main.ts`)

**Files:**
- Modify: `src/main.ts` (replace sample-plugin contents entirely)

**Interfaces:**
- Consumes: everything from Tasks 1–12.
- Produces: the `TodoSyncPlugin` default export that `settings.ts` (Task 12) types against.

This is lifecycle wiring, not business logic — the business logic it calls (`SyncEngine.runPollCycle`, `MsalDeviceCodeAuth`, `TodoClient`) is already tested. Covered by Task 14's manual verification, same as the rest of the Obsidian-runtime glue.

- [ ] **Step 1: Replace `src/main.ts`**

```ts
import { Notice, Plugin } from 'obsidian';
import { DEFAULT_SETTINGS, TodoSyncSettings, TodoSyncSettingTab } from './settings';
import { createVaultAdapter } from './obsidian-tasks/vault-adapter';
import { createCachePlugin, hasStoredSession } from './auth/token-storage';
import { MsalDeviceCodeAuth } from './auth/msal-device-code';
import { TodoClient } from './todo-api/client';
import { SyncEngine } from './sync/engine';
import { loadSyncData, serializeSyncData, type SyncData } from './sync/state-store';
import { DeviceCodeModal } from './ui/device-code-modal';
import { RetryAfterError } from './todo-api/types';

interface PersistedData {
	settings?: Partial<TodoSyncSettings>;
	sync?: unknown;
}

export default class TodoSyncPlugin extends Plugin {
	settings!: TodoSyncSettings;
	private syncData!: SyncData;
	private auth!: MsalDeviceCodeAuth;
	private todoClient!: TodoClient;
	private engine!: SyncEngine;
	private pollIntervalId: number | null = null;

	async onload() {
		await this.loadPersistedData();

		this.auth = new MsalDeviceCodeAuth(createCachePlugin(this.app.secretStorage));
		this.todoClient = new TodoClient(() => this.getAccessToken());
		this.engine = new SyncEngine({
			todo: this.todoClient,
			vault: createVaultAdapter(this.app),
			now: () => new Date(),
			reminderTime: this.settings.reminderTime,
			listName: this.settings.listName,
			log: (message) => console.log(`[todo-sync] ${message}`),
		});

		this.addSettingTab(new TodoSyncSettingTab(this.app, this));

		this.addCommand({
			id: 'sign-in-to-microsoft',
			name: 'Sign in to Microsoft',
			callback: () => this.signIn(),
		});
		this.addCommand({
			id: 'sign-out-of-microsoft',
			name: 'Sign out',
			callback: () => this.signOut(),
		});
		this.addCommand({
			id: 'sync-now',
			name: 'Sync now',
			callback: () => this.runSyncCycle(),
		});

		if (hasStoredSession(this.app.secretStorage)) {
			this.startPolling();
		}
	}

	onunload() {
		this.stopPolling();
	}

	async loadPersistedData(): Promise<void> {
		const raw = ((await this.loadData()) ?? {}) as PersistedData;
		this.settings = { ...DEFAULT_SETTINGS, ...raw.settings };
		this.syncData = loadSyncData(raw.sync);
	}

	async saveSettings(): Promise<void> {
		await this.persist();
	}

	private async persist(): Promise<void> {
		const data: PersistedData = {
			settings: this.settings,
			sync: serializeSyncData(this.syncData),
		};
		await this.saveData(data);
	}

	private async getAccessToken(): Promise<string> {
		const token = await this.auth.getAccessTokenSilent();
		if (!token) {
			this.stopPolling();
			new Notice('Microsoft sign-in expired. Run "Sign in to Microsoft" to reconnect.');
			throw new Error('No valid access token; re-authentication required');
		}
		return token;
	}

	async signIn(): Promise<void> {
		const modal = new DeviceCodeModal(this.app);
		modal.open();
		try {
			await this.auth.signIn((info) => modal.showDeviceCode(info));
			this.settings.signedInAccountLabel = 'Microsoft account';
			await this.saveSettings();
			modal.showSuccess(this.settings.signedInAccountLabel);
			this.startPolling();
			await this.runSyncCycle();
		} catch (error) {
			modal.showError(error instanceof Error ? error.message : String(error));
		}
	}

	async signOut(): Promise<void> {
		this.stopPolling();
		await this.auth.signOut();
		this.settings.signedInAccountLabel = null;
		await this.saveSettings();
		new Notice('Signed out of Microsoft To Do sync.');
	}

	async handleSignInOutFromSettings(): Promise<void> {
		if (this.settings.signedInAccountLabel) {
			await this.signOut();
		} else {
			await this.signIn();
		}
		this.app.workspace.trigger('layout-change');
	}

	startPolling(): void {
		this.stopPolling();
		const intervalMs = this.settings.pollIntervalMinutes * 60 * 1000;
		this.pollIntervalId = window.setInterval(() => this.runSyncCycle(), intervalMs);
		this.registerInterval(this.pollIntervalId);
	}

	restartPolling(): void {
		if (this.pollIntervalId !== null) this.startPolling();
	}

	private stopPolling(): void {
		if (this.pollIntervalId !== null) {
			window.clearInterval(this.pollIntervalId);
			this.pollIntervalId = null;
		}
	}

	async runSyncCycle(): Promise<void> {
		try {
			this.syncData = await this.engine.runPollCycle(this.syncData);
			await this.persist();
		} catch (error) {
			if (error instanceof RetryAfterError) {
				console.warn(`[todo-sync] throttled, backing off ${error.retryAfterSeconds}s`);
				return;
			}
			console.error('[todo-sync] sync cycle failed', error);
		}
	}
}
```

- [ ] **Step 2: Run the full test suite**

Run: `npm test`
Expected: all suites still pass (this task adds no new automated tests — see rationale above).

- [ ] **Step 3: Type-check and build**

Run: `npx tsc -noEmit -skipLibCheck`
Expected: no errors. Fix any signature drift between `settings.ts` (Task 12) and the methods now defined on `TodoSyncPlugin` (`handleSignInOutFromSettings`, `restartPolling`, `saveSettings`) before proceeding.

Run: `npm run build`
Expected: succeeds, produces `main.js`.

- [ ] **Step 4: Commit**

```bash
git add src/main.ts
git commit -m "Wire auth, Graph client, and sync engine into plugin lifecycle"
```

---

## Task 14: Manual end-to-end verification

**Files:** none (verification only).

This task exists because CLAUDE.md requires runtime behaviour (UI, network, vault writes) to be exercised, not just type-checked — the automated suite in Tasks 0–10 covers all business logic, but sign-in, settings rendering, and actual vault mutation via the real Obsidian API have not been run for real yet.

- [ ] **Step 1: Register the real Entra ID app**

In Entra ID admin center: create a new app registration, platform "Mobile and desktop applications" with the `https://login.microsoftonline.com/common/oauth2/nativeclient` redirect URI, public client flows enabled, delegated API permission `Tasks.ReadWrite` granted. Copy the Application (client) ID into `src/auth/config.ts`'s `CLIENT_ID`.

- [ ] **Step 2: Build and install into a test vault**

```bash
npm run build
```

Copy `main.js`, `manifest.json`, `styles.css` into `<TestVault>/.obsidian/plugins/obsidian-todo-sync/`.

- [ ] **Step 3: Enable and sign in**

Reload Obsidian, enable "Microsoft To Do Sync" in **Settings → Community plugins**. Run the "Sign in to Microsoft" command. Confirm the modal shows a device code and verification URL, complete sign-in in a browser, confirm the modal shows "Signed in as...".

- [ ] **Step 4: Verify Obsidian -> To Do**

Add `- [ ] Manual test task 📅 2026-10-15` to a note. Run "Sync now". Confirm: the line gains a `%%todo:...%%` marker, and the task appears in the "Obsidian" To Do list with the due date and an 08:00 reminder set.

- [ ] **Step 5: Verify To Do -> Obsidian (completion)**

Mark that task complete in the To Do app. Run "Sync now". Confirm the Obsidian line becomes `- [x] ... ✅ <today> %%todo:...%%`.

- [ ] **Step 6: Verify new-from-remote**

Create a new task directly in the "Obsidian" To Do list with a due date, no corresponding Obsidian line. Ensure the current ISO-week weekly note exists with an `### Other Action Items` heading. Run "Sync now". Confirm a new line appears under that heading with the marker.

- [ ] **Step 7: Verify pending queue when the weekly note is missing**

Delete (or don't create) the current week's weekly note. Create another new To Do task. Run "Sync now" — confirm no error and no orphaned line anywhere. Create the weekly note (matching `Calendar/Weekly/YYYY-Www.md` with the heading). Run "Sync now" again. Confirm the queued task now appears.

- [ ] **Step 8: Verify recurring tasks are skipped**

Add `- [ ] Recurring test 📅 2026-10-15 🔁 every week` to a note. Run "Sync now". Confirm it is never pushed to To Do (no marker added, nothing created remotely).

- [ ] **Step 9: Verify sign-out**

Run "Sign out". Confirm settings show "Not signed in" and polling stops (no further Notices/log lines after the interval elapses).

- [ ] **Step 10: Run lint**

Run: `npm run lint`
Expected: no errors (fix any `eslint-plugin-obsidianmd` findings before considering this plan complete).

- [ ] **Step 11: Final commit**

If Step 10 required fixes:

```bash
git add -A
git commit -m "Fix lint findings from manual verification pass"
```

---

## Self-Review Notes

- **Spec coverage:** every "In scope" bullet has a task — new-local-push (Task 10), remote-completion (Task 10), local-completion (Task 10, `applyLocalToRemote`), edits both ways (Task 10), new-from-remote + pending queue (Tasks 7, 10), recurrence skip (Task 2 detection + Task 10 skip), desktop-only (Task 0 manifest). Auth (Tasks 5–6), token storage (Task 5), sync-identity marker (Task 2), settings (Task 12), commands (Task 13), error handling — auth failure/expired token (Task 13 `getAccessToken`), network failure (Task 13 `runSyncCycle` catch-all), 429 back-off (Task 4 `RetryAfterError` + Task 13 catch), conflict last-write-wins (Task 10).
- **Out of scope confirmed excluded:** no multi-list/multi-vault settings, no mobile handling beyond `isDesktopOnly: true`, no subtask/attachment parsing.
- **Type consistency checked:** `ParsedTaskLine` (Task 1) fields match `parser.ts` (Task 2) and every construction site in `engine.ts` (Task 10). `SyncEngineDeps.todo` in Task 10 is a `Pick<TodoClient, ...>` matching the four methods `TodoClient` (Task 4) actually implements. `VaultAdapter` methods used in Task 10 (`listMarkdownFiles`, `read`, `update`, `getFile`) match Task 9's interface exactly. `PendingRemoteTask` shape is identical across Task 7, Task 8's `SyncData.pending`, and Task 10's usage.
- **Open items from the spec not blocking this plan** (carried into Task 14 manual verification, per the spec's own framing): whether `reminderDateTime` reliably fires a client notification in Frank's tenant; whether the tenant's consent policy allows user consent for `Tasks.ReadWrite`; whether Graph enforces `@odata.etag`/`If-Match` on `PATCH`.
