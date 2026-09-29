import type { ParsedTaskLine } from '../types';

const CHECKBOX_RE = /^-\s\[([ xX])\]\s(.*)$/;
const DUE_DATE_RE = /📅\s*(\d{4}-\d{2}-\d{2})/;
const DONE_DATE_RE = /✅\s*(\d{4}-\d{2}-\d{2})/;
const NO_DUE_RE = /\bNO DUE DATE\b/;
const DANGLING_DUE_RE = /\s*📅\s*$/;
const RECURRENCE_RE = /🔁[^\n]*$/;
export const MARKER_RE = /%%todo:([^%]+)%%/;

export function parseTaskLine(line: string): ParsedTaskLine | null {
	const indentMatch = /^(\s*)/.exec(line);
	const indent = indentMatch![1]!;
	const checkboxMatch = CHECKBOX_RE.exec(line.trim());
	if (!checkboxMatch) return null;

	const [, mark, rest] = checkboxMatch;
	const dueMatch = DUE_DATE_RE.exec(rest!);
	const markerMatch = MARKER_RE.exec(rest!);
	// A marker alone qualifies: synced tasks without a due date have no 📅.
	if (!dueMatch && !markerMatch) return null;

	const doneMatch = DONE_DATE_RE.exec(rest!);
	const recurring = RECURRENCE_RE.test(rest!);

	const title = rest!
		.replace(MARKER_RE, '')
		.replace(RECURRENCE_RE, '')
		.replace(DONE_DATE_RE, '')
		.replace(DUE_DATE_RE, '')
		.replace(NO_DUE_RE, '')
		.replace(DANGLING_DUE_RE, '')
		.trim();

	return {
		indent,
		checked: mark!.toLowerCase() === 'x',
		title,
		dueDate: dueMatch ? dueMatch[1]! : null,
		doneDate: doneMatch ? doneMatch[1]! : null,
		recurring,
		todoId: markerMatch ? markerMatch[1]! : null,
	};
}

export function renderTaskLine(task: ParsedTaskLine): string {
	const mark = task.checked ? 'x' : ' ';
	// Marker goes before the emoji metadata: the Tasks plugin only reads
	// 📅/✅ when nothing but metadata follows, so a trailing marker hides the task.
	let line = `${task.indent}- [${mark}] ${task.title}`;
	if (task.todoId) line += ` %%todo:${task.todoId}%%`;
	line += task.dueDate ? ` 📅 ${task.dueDate}` : ' NO DUE DATE';
	if (task.doneDate) line += ` ✅ ${task.doneDate}`;
	return line;
}
