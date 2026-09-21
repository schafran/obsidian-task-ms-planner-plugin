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
