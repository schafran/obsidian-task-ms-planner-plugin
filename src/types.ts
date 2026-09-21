export interface ParsedTaskLine {
	indent: string;
	checked: boolean;
	title: string;
	dueDate: string | null; // YYYY-MM-DD
	doneDate: string | null; // YYYY-MM-DD
	recurring: boolean;
	todoId: string | null;
}
