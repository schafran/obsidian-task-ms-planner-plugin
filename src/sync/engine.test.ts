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
