import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';
import { Paragraph } from '../../mod.ts';
import { Archive } from '../classes/Archive.ts';
import type { ComponentContext } from '../classes/Component.ts';
import { create, serialize } from '../utilities/dom.ts';
import { NamespaceUri } from '../utilities/namespaces.ts';

import { Text } from './Text.ts';

describe('Deletion', () => {
	const date = new Date();

	const emptyContext: ComponentContext = {
		archive: new Archive(),
		relationships: null,
	};

	describe('Inserted run content', () => {});
	describe('Inserted numbering properties', () => {});
	describe('Inserted paragraph', () => {
		const deletedParagraphNode = create(
			`<w:p xmlns:w="${NamespaceUri.w}">
				<w:pPr>
					<w:rPr>
						<w:del w:id="1" w:author="Luis" w:date="${date.toISOString()}" />
					</w:rPr>
				</w:pPr>
				<w:r>
					<w:t>This is paragraph one.</w:t>
				</w:r>
			</w:p>
			`,
			emptyContext
		);

		const deletedParagraphAsProp = new Paragraph(
			{ pilcrow: { deletion: { author: 'Luis', date: date, id: 1 } } },
			new Text({}, 'This is paragraph one.')
		);

		const deletedParagraphAsNode = Paragraph.fromNode(
			deletedParagraphNode,
			emptyContext
		);

		it('Paragraph node has expected deletion objects', () => {
			expect(deletedParagraphAsNode.props.pilcrow?.deletion).toEqual(
				deletedParagraphAsProp.props.pilcrow?.deletion
			);
		});

		it('serializes and deserialized correctly', async () => {
			expect(serialize(await deletedParagraphAsProp.toNode([]))).toEqual(
				serialize(
					create(
						`<p xmlns="${NamespaceUri.w}">
							<pPr>
								<rPr>
									<del xmlns:ns1="${
										NamespaceUri.w
									}" ns1:id="1" ns1:author="Luis" ns1:date="${date.toISOString()}" />
								</rPr>
							</pPr>
							<r>
								<t xml:space="preserve">This is paragraph one.</t>
							</r>
						</p>`
					)
				)
			);
		});
	});
});
