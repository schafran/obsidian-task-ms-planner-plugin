import type { ParsedTaskLine } from '../types';

const CHECKBOX_RE = /^-\s\[([ xX])\]\s(.*)$/;
const DUE_DATE_RE = /📅\s*(\d{4}-\d{2}-\d{2})/;
const DONE_DATE_RE = /✅\s*(\d{4}-\d{2}-\d{2})/;
const RECURRENCE_RE = /🔁[^\n]*$/;
const MARKER_RE = /%%todo:([^%]+)%%/;

export function parseTaskLine(line: string): ParsedTaskLine | null {
	const checkboxMatch = CHECKBOX_RE.exec(line.trim());
	if (!checkboxMatch) return null;

	const [, mark, rest] = checkboxMatch;
	const dueMatch = DUE_DATE_RE.exec(rest);
	if (!dueMatch) return null;

	const doneMatch = DONE_DATE_RE.exec(rest);
	const markerMatch = MARKER_RE.exec(rest);
	const recurring = RECURRENCE_RE.test(rest);

	const title = rest
		.replace(MARKER_RE, '')
		.replace(RECURRENCE_RE, '')
		.replace(DONE_DATE_RE, '')
		.replace(DUE_DATE_RE, '')
		.trim();

	return {
		checked: mark.toLowerCase() === 'x',
		title,
		dueDate: dueMatch[1],
		doneDate: doneMatch ? doneMatch[1] : null,
		recurring,
		todoId: markerMatch ? markerMatch[1] : null,
	};
}

export function renderTaskLine(task: ParsedTaskLine): string {
	const mark = task.checked ? 'x' : ' ';
	let line = `- [${mark}] ${task.title} 📅 ${task.dueDate}`;
	if (task.doneDate) line += ` ✅ ${task.doneDate}`;
	if (task.todoId) line += ` %%todo:${task.todoId}%%`;
	return line;
}
