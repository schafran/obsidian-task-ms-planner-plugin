import { describe, expect, it } from 'vitest';
import { PendingQueue } from './pending-queue';

describe('PendingQueue', () => {
	it('starts empty', () => {
		expect(new PendingQueue().all).toEqual([]);
	});

	it('adds items and exposes them via all', () => {
		const queue = new PendingQueue();
		queue.add({ todoId: 't1', title: 'Buy milk', dueDate: '2026-10-01' });
		expect(queue.all).toEqual([{ todoId: 't1', title: 'Buy milk', dueDate: '2026-10-01' }]);
	});

	it('removes an item by todoId', () => {
		const queue = new PendingQueue();
		queue.add({ todoId: 't1', title: 'Buy milk', dueDate: '2026-10-01' });
		queue.add({ todoId: 't2', title: 'Water plants', dueDate: '2026-10-02' });
		queue.remove('t1');
		expect(queue.all.map((t) => t.todoId)).toEqual(['t2']);
	});

	it('round-trips through toJSON/fromJSON', () => {
		const queue = new PendingQueue();
		queue.add({ todoId: 't1', title: 'Buy milk', dueDate: '2026-10-01' });
		const restored = PendingQueue.fromJSON(queue.toJSON());
		expect(restored.all).toEqual(queue.all);
	});

	it('fromJSON tolerates undefined (first run, no persisted data yet)', () => {
		expect(PendingQueue.fromJSON(undefined).all).toEqual([]);
	});
});
