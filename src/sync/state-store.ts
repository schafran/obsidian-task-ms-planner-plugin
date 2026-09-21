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
