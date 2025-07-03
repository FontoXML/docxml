import { describe, it } from 'std/testing/bdd';

import { MoveRangeEnd } from '@fontoxml/docxml';
import { expect } from 'std/expect';
import { Archive } from '../classes/Archive.ts';
import type { ComponentContext } from '../classes/Component.ts';
import { create, serialize } from '../utilities/dom.ts';
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

	it('deserialize nodes correctly', () => {
		const deserializedTo = newMoveToRangeStart.toNode();
		const targetTo = create(
			`<moveToRangeStart name="Move_1" id="0" date="${date.toISOString()}" author="Gabe"/>`
		);
		expect(deserializedTo).toEqual(targetTo);
	});
});
