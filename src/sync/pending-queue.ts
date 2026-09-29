export interface PendingRemoteTask {
	todoId: string;
	title: string;
	dueDate: string | null; // legacy entries may hold ''
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
