import { RangeSetBuilder } from '@codemirror/state';
import {
	Decoration,
	EditorView,
	ViewPlugin,
	WidgetType,
	type DecorationSet,
	type ViewUpdate,
} from '@codemirror/view';
import { editorLivePreviewField } from 'obsidian';
import { findMarkerRanges } from './marker-ranges';

class MarkerWidget extends WidgetType {
	eq(): boolean {
		return true;
	}

	toDOM(): HTMLElement {
		const el = createSpan({ cls: 'todo-sync-marker', text: '🔗' });
		el.title = 'Synced with Microsoft To Do';
		return el;
	}
}

const widget = Decoration.replace({ widget: new MarkerWidget() });

function buildDecorations(view: EditorView): DecorationSet {
	const builder = new RangeSetBuilder<Decoration>();
	if (!view.state.field(editorLivePreviewField, false)) return builder.finish();

	const { doc, selection } = view.state;
	let lastLine = -1;
	for (const { from, to } of view.visibleRanges) {
		for (let pos = from; pos <= to; ) {
			const line = doc.lineAt(pos);
			pos = line.to + 1;
			if (line.number === lastLine) continue;
			lastLine = line.number;
			for (const range of findMarkerRanges(line.text, line.from)) {
				const touched = selection.ranges.some((s) => s.from <= range.to && s.to >= range.from);
				if (!touched) builder.add(range.from, range.to, widget);
			}
		}
	}
	return builder.finish();
}

const hideMarkerPlugin = ViewPlugin.fromClass(
	class {
		decorations: DecorationSet;

		constructor(view: EditorView) {
			this.decorations = buildDecorations(view);
		}

		update(update: ViewUpdate): void {
			if (update.docChanged || update.selectionSet || update.viewportChanged || update.geometryChanged) {
				this.decorations = buildDecorations(update.view);
			}
		}
	},
	{
		decorations: (v) => v.decorations,
		provide: (plugin) =>
			EditorView.atomicRanges.of((view) => view.plugin(plugin)?.decorations ?? Decoration.none),
	},
);

export const hideMarkerExtension = [hideMarkerPlugin];
