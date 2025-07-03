import { describe, it } from 'std/testing/bdd';

import { Move } from '@fontoxml/docxml';
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

	const moveFromNode = create(
		`<w:moveFrom xmlns:w="${
			NamespaceUri.w
		}" w:id="0" w:author="Gabe" w:date="${date.toISOString()}">
			<w:p>
				<w:r>
					<w:t>This is paragraph text</w:t>
				</w:r>
			</w:p>
		
		</w:moveFrom>`,
		emptyContext
	);

	const moveToNode = create(
		`<w:moveTo xmlns:w="${
			NamespaceUri.w
		}" w:id="0" w:author="Gabe" w:date="${date.toISOString()}">
			<w:p>
				<w:r>
					<w:t>This is paragraph text</w:t>
				</w:r>
			</w:p>
		
		</w:moveTo>`,
		emptyContext
	);

	const newMoveFrom = Move.fromNode(moveFromNode, emptyContext);
	const newMoveTo = Move.fromNode(moveToNode, emptyContext);
	console.log(newMoveFrom);
	console.log(newMoveTo);
});
