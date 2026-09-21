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
