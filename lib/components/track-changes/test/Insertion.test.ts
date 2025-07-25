import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';
import { Paragraph, Text } from '../../../../mod.ts';
import { Archive } from '../../../classes/src/Archive.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { create, serialize } from '../../../utilities/src/dom.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';

describe('Insertion', () => {
	const date = new Date();

	const emptyContext: ComponentContext = {
		archive: new Archive(),
		relationships: null,
	};

	describe('Inserted paragraph', () => {
		const insertedParagraphNode = create(
			`<w:p xmlns:w="${NamespaceUri.w}">
				<w:pPr>
					<w:rPr>
						<w:ins w:id="1" w:author="Luis" w:date="${date.toISOString()}" />
					</w:rPr>
				</w:pPr>
				<w:r>
					<w:t>This is paragraph one.</w:t>
				</w:r>
			</w:p>
			`,
			emptyContext
		);

		const insertedParagraphAsProp = new Paragraph(
			{ pilcrow: { insertion: { author: 'Luis', date: date, id: 1 } } },
			new Text({}, 'This is paragraph one.')
		);

		const insertedParagraphAsNode = Paragraph.fromNode(
			insertedParagraphNode,
			emptyContext
		);

		it('Paragraph node has expected insertion objects', () => {
			expect(insertedParagraphAsNode.props.pilcrow?.insertion).toEqual(
				insertedParagraphAsProp.props.pilcrow?.insertion
			);
		});

		it('serializes and deserialized correctly', async () => {
			expect(serialize(await insertedParagraphAsProp.toNode([]))).toEqual(
				serialize(
					create(
						`<p xmlns="${NamespaceUri.w}">
							<pPr>
								<rPr>
									<ins xmlns:ns1="${
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
