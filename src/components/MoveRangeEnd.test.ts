import { describe, it } from 'std/testing/bdd';

import { Archive } from '../classes/Archive.ts';
import type { ComponentContext } from '../classes/Component.ts';
import { MoveRangeEnd } from './MoveRangeEnd.ts';

describe('Move range start for to and from elements...', () => {
	const date = new Date();

	const emptyContext: ComponentContext = {
		archive: new Archive(),
		relationships: null,
	};

	const newMoveToRangeEnd = new MoveRangeEnd({
		type: 'to',
		id: 0,
	});

	const newMoveFromRangeEnd = new MoveRangeEnd({
		type: 'from',
		id: 0,
	});

	console.log(newMoveToRangeEnd);
	console.log(newMoveFromRangeEnd);
});
