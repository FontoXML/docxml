import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Move } from './Move.ts';
import { Paragraph } from './Paragraph.ts';
import { Text } from './Text.ts';

import { Archive } from '../classes/Archive.ts';
import type { ComponentContext } from '../classes/Component.ts';
import { create, serialize } from '../utilities/dom.ts';
import { NamespaceUri } from '../utilities/namespaces.ts';

describe('Move content in track changes...', () => {
	const date = new Date();

	const emptyContext: ComponentContext = {
		archive: new Archive(),
		relationships: null,
	};

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

	const moveFromNode = create(
		`<w:moveFrom xmlns:w="${
			NamespaceUri.w
		}" w:id="1" w:author="Angel" w:date="${date.toISOString()}">
			<w:p>
				<w:r>
					<w:t>This is a moveFrom node.</w:t>
				</w:r>
			</w:p>
		
		</w:moveFrom>`,
		emptyContext
	);

	const moveToObject = new Move(
		{
			id: 0,
			date: date,
			author: 'Gabe',
			type: 'to',
		},
		new Paragraph({ style: null }, new Text({}, 'This is paragraph text'))
	);

	const moveFromObject = new Move(
		{
			id: 1,
			date: date,
			author: 'Angel',
			type: 'from',
		},
		new Paragraph({ style: null }, new Text({}, 'This is a moveFrom node.'))
	);

	const newMoveTo = Move.fromNode(moveToNode, emptyContext);
	const newMoveFrom = Move.fromNode(moveFromNode, emptyContext);

	it('turns node into correct Move objects', () => {
		expect(newMoveTo).toEqual(moveToObject);
		expect(newMoveFrom).toEqual(moveFromObject);
	});

	it('turns Move object into the correct node', async () => {
		expect(serialize(await moveToObject.toNode([]))).toEqual(
			serialize(
				create(
					`<p xmlns="${NamespaceUri.w}" xmlns:ns1="${
						NamespaceUri.w
					}"  >
						<pPr><moveTo ns1:id="0" ns1:date="${date.toISOString()}" ns1:author="Gabe"/></pPr>
						<r>
							<t xml:space="preserve" >This is paragraph text</t>
						</r>
					</p>
					`
				)
			)
		);
	});

	it('turns Move object into the correct node', async () => {
		expect(serialize(await moveFromObject.toNode([]))).toEqual(
			serialize(
				create(
					`<p xmlns="${NamespaceUri.w}" xmlns:ns1="${
						NamespaceUri.w
					}" >
						<pPr>
							<moveFrom  ns1:id="1" ns1:date="${date.toISOString()}" ns1:author="Angel" /> 
						</pPr>
						<r>
							<t xml:space="preserve" >This is a moveFrom node.</t>
						</r>
					</p>`
				)
			)
		);
	});
});
