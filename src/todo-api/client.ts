import { requestUrl, type RequestUrlParam } from 'obsidian';
import { RetryAfterError, type DeltaResult, type NewTaskInput, type TodoList, type TodoTask } from './types';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

interface RequestInit {
	method?: string;
	headers?: Record<string, string>;
	body?: string;
}

export class TodoClient {
	constructor(private readonly getAccessToken: () => Promise<string>) {}

	private async request<T>(pathOrUrl: string, init: RequestInit = {}): Promise<T> {
		const token = await this.getAccessToken();
		const url = pathOrUrl.startsWith('http') ? pathOrUrl : `${GRAPH_BASE}${pathOrUrl}`;
		const params: RequestUrlParam = {
			url,
			method: init.method,
			headers: {
				...init.headers,
				Authorization: `Bearer ${token}`,
				'Content-Type': 'application/json',
			},
			body: init.body,
			throw: false,
		};
		const res = await requestUrl(params);

		if (res.status === 429) {
			const retryAfterHeader =
				res.headers['Retry-After'] ?? res.headers['retry-after'] ?? '60';
			throw new RetryAfterError(Number(retryAfterHeader));
		}
		if (res.status < 200 || res.status >= 300) {
			throw new Error(`Graph request failed: ${res.status} ${res.text}`);
		}
		if (res.status === 204) return undefined as T;
		return res.json as T;
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

		const tasks = data.value.filter((t) => !('@removed' in t));

		if (data['@odata.nextLink']) {
			const next = await this.fetchDelta(listId, data['@odata.nextLink']);
			return { tasks: [...tasks, ...next.tasks], deltaLink: next.deltaLink };
		}
		return { tasks, deltaLink: data['@odata.deltaLink'] ?? '' };
	}
}
