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
		fetchMock
			.mockResolvedValueOnce(jsonResponse({}, 429, { 'Retry-After': '30' }))
			.mockResolvedValueOnce(jsonResponse({}, 429, { 'Retry-After': '30' }));
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
