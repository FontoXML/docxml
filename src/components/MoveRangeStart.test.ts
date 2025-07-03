import { describe, it } from 'std/testing/bdd';

import { Archive } from '../classes/Archive.ts';
import type { ComponentContext } from '../classes/Component.ts';
import { MoveRangeStart } from './MoveRangeStart.ts';

describe('Move range start for to and from elements...', () => {
	const date = new Date();

	const newMoveToRangeStart = new MoveRangeStart({
		type: 'to',
		name: 'Move_1',
		id: 0,
		date: date,
		author: 'Gabe',
	});

	const newMoveFromRangeStart = new MoveRangeStart({
		type: 'from',
		name: 'Move_1',
		id: 0,
		date: date,
		author: 'Angel',
	});

	console.log(newMoveToRangeStart);
	console.log(newMoveFromRangeStart);
});
