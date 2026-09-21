import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from 'vitest';
import { requestUrl, type RequestUrlParam, type RequestUrlResponse } from 'obsidian';
import { TodoClient } from './client';
import { RetryAfterError } from './types';

function urlResponse(
	body: unknown,
	status = 200,
	headers: Record<string, string> = {},
): RequestUrlResponse {
	return {
		status,
		headers,
		json: body,
		text: JSON.stringify(body),
		arrayBuffer: new ArrayBuffer(0),
	};
}

describe('TodoClient', () => {
	const getAccessToken = vi.fn(async () => 'fake-token');
	let requestUrlMock: Mock<[params: RequestUrlParam], Promise<RequestUrlResponse>>;

	beforeEach(() => {
		requestUrlMock = requestUrl as unknown as Mock<
			[params: RequestUrlParam],
			Promise<RequestUrlResponse>
		>;
		requestUrlMock.mockReset();
	});

	afterEach(() => {
		getAccessToken.mockClear();
	});

	it('ensureList returns an existing list by display name without creating one', async () => {
		requestUrlMock.mockResolvedValueOnce(
			urlResponse({ value: [{ id: '1', displayName: 'Obsidian' }] }),
		);
		const client = new TodoClient(getAccessToken);
		const list = await client.ensureList('Obsidian');
		expect(list).toEqual({ id: '1', displayName: 'Obsidian' });
		expect(requestUrlMock).toHaveBeenCalledTimes(1);
	});

	it('ensureList creates the list when none matches', async () => {
		requestUrlMock
			.mockResolvedValueOnce(urlResponse({ value: [] }))
			.mockResolvedValueOnce(urlResponse({ id: '2', displayName: 'Obsidian' }));
		const client = new TodoClient(getAccessToken);
		const list = await client.ensureList('Obsidian');
		expect(list).toEqual({ id: '2', displayName: 'Obsidian' });
		expect(requestUrlMock).toHaveBeenCalledTimes(2);
		const [createParams] = requestUrlMock.mock.calls[1]!;
		expect(JSON.parse(createParams.body as string)).toEqual({ displayName: 'Obsidian' });
	});

	it('createTask sends title, due date, and reminder', async () => {
		requestUrlMock.mockResolvedValueOnce(urlResponse({ id: 't1' }));
		const client = new TodoClient(getAccessToken);
		await client.createTask('list1', {
			title: 'Renew passport',
			dueDate: '2026-10-01',
			reminderTime: '08:00',
		});
		const [params] = requestUrlMock.mock.calls[0]!;
		expect(params.url).toBe('https://graph.microsoft.com/v1.0/me/todo/lists/list1/tasks');
		const body = JSON.parse(params.body as string) as {
			title: string;
			dueDateTime: unknown;
			reminderDateTime: unknown;
			isReminderOn: boolean;
		};
		expect(body.title).toBe('Renew passport');
		expect(body.dueDateTime).toEqual({ dateTime: '2026-10-01T00:00:00', timeZone: 'UTC' });
		expect(body.reminderDateTime).toEqual({
			dateTime: '2026-10-01T08:00:00',
			timeZone: 'UTC',
		});
		expect(body.isReminderOn).toBe(true);
	});

	it('updateTask PATCHes only the given fields', async () => {
		requestUrlMock.mockResolvedValueOnce(urlResponse({ id: 't1', status: 'completed' }));
		const client = new TodoClient(getAccessToken);
		await client.updateTask('list1', 't1', { status: 'completed' });
		const [params] = requestUrlMock.mock.calls[0]!;
		expect(params.url).toBe('https://graph.microsoft.com/v1.0/me/todo/lists/list1/tasks/t1');
		expect(params.method).toBe('PATCH');
		expect(JSON.parse(params.body as string)).toEqual({ status: 'completed' });
	});

	it('fetchDelta follows nextLink pages and returns the final deltaLink', async () => {
		requestUrlMock
			.mockResolvedValueOnce(
				urlResponse({
					value: [{ id: 'a' }],
					'@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/todo/lists/list1/tasks/delta?$skiptoken=1',
				}),
			)
			.mockResolvedValueOnce(
				urlResponse({ value: [{ id: 'b' }], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/final' }),
			);
		const client = new TodoClient(getAccessToken);
		const result = await client.fetchDelta('list1');
		expect(result.tasks.map((t) => t.id)).toEqual(['a', 'b']);
		expect(result.deltaLink).toBe('https://graph.microsoft.com/v1.0/final');
	});

	it('filters @removed tombstones out of the returned tasks', async () => {
		requestUrlMock.mockResolvedValueOnce(
			urlResponse({
				value: [{ id: 'a', title: 'Normal task' }, { id: 'x', '@removed': { reason: 'deleted' } }],
				'@odata.deltaLink': 'https://graph.microsoft.com/v1.0/final',
			}),
		);
		const client = new TodoClient(getAccessToken);
		const result = await client.fetchDelta('list1');
		expect(result.tasks.map((t) => t.id)).toEqual(['a']);
	});

	it('throws RetryAfterError on 429 with the Retry-After value', async () => {
		requestUrlMock
			.mockResolvedValueOnce(urlResponse({}, 429, { 'Retry-After': '30' }))
			.mockResolvedValueOnce(urlResponse({}, 429, { 'Retry-After': '30' }));
		const client = new TodoClient(getAccessToken);
		await expect(client.fetchDelta('list1')).rejects.toBeInstanceOf(RetryAfterError);
		try {
			await client.fetchDelta('list1');
		} catch (e) {
			expect((e as InstanceType<typeof RetryAfterError>).retryAfterSeconds).toBe(30);
		}
	});

	it('throws a descriptive error on other non-2xx responses', async () => {
		requestUrlMock.mockResolvedValueOnce(urlResponse({ error: 'nope' }, 500));
		const client = new TodoClient(getAccessToken);
		await expect(client.fetchDelta('list1')).rejects.toThrow(/500/);
	});
});
