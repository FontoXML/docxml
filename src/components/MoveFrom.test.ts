import { describe, it } from 'std/testing/bdd';

import { MoveFrom } from '@fontoxml/docxml';
import { Archive } from '../classes/Archive.ts';
import type { ComponentContext } from '../classes/Component.ts';
import { create } from '../utilities/dom.ts';
import { NamespaceUri } from '../utilities/namespaces.ts';

describe('Move content in track changes...', () => {
	const date = new Date();

	const emptyContext: ComponentContext = {
		archive: new Archive(),
		relationships: null,
	};

	const moveNode = create(
		`
        <w:moveFromRangeStart xmlns:w="${
			NamespaceUri.w
		}" w:id="0" w:name="Move_0" w:author="Gabe" w:date="${date.toISOString()}" />
        <w:moveFromRangeEnd xmlns:w="${NamespaceUri.w}" w:id="0" />`,
		emptyContext
	);

	const newMove = MoveFrom.fromNode(moveNode, emptyContext);
	console.log(newMove);
});
