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
