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
			if (!state) continue;

			// A task absent from the delta response means it has not changed remotely
			// since the last delta cursor (Graph delta semantics: unchanged items are
			// omitted). Only treat it as remote-changed when it's actually present.
			const remoteChanged = remote
				? new Date(remote.lastModifiedDateTime).getTime() > state.lastSyncedAtMs
				: false;
			const localChanged = located.file.mtimeMs > state.lastSyncedAtMs;

			if (remoteChanged && !localChanged) {
				await this.applyRemoteToLocal(located, remote!);
			} else if (localChanged && !remoteChanged) {
				await this.applyLocalToRemote(list.id, located, todo);
			} else if (remoteChanged && localChanged) {
				if (new Date(remote!.lastModifiedDateTime).getTime() >= located.file.mtimeMs) {
					await this.applyRemoteToLocal(located, remote!);
				} else {
					await this.applyLocalToRemote(list.id, located, todo);
				}
			}

			taskStates[todoId] = {
				lastKnownRemoteModified: remote ? remote.lastModifiedDateTime : state.lastKnownRemoteModified,
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
