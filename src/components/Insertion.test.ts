import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';
import { Paragraph } from '../../mod.ts';
import { Archive } from '../classes/Archive.ts';
import type { ComponentContext } from '../classes/Component.ts';
import { create, serialize } from '../utilities/dom.ts';
import { NamespaceUri } from '../utilities/namespaces.ts';
import { Cell } from './Cell.ts';
import { Row } from './Row.ts';
import { Table } from './Table.ts';
import { Text } from './Text.ts';

describe('Insertion', () => {
	const date = new Date();

	const emptyContext: ComponentContext = {
		archive: new Archive(),
		relationships: null,
	};

	describe('Inserted run content', () => {});
	describe('Inserted numbering properties', () => {});
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
	describe('Inserted table row', () => {
		const insertedRowNode = create(
			`
            <w:tr xmlns:w="${NamespaceUri.w}">
                <w:trPr>
                    <w:tblHeader/>
                    <w:cantSplit/>
                    <w:tblCellSpacing w:w="1701" w:type="dxa"/>
                    <w:ins w:id="1" w:author="Luis" w:date="${date.toISOString()}"/>
                </w:trPr>
                <w:tc>
                <w:tcPr/>
                <w:p>
                    <w:r>
                        <w:t xml:space="preserve"> it is time</w:t>
                    </w:r>
                </w:p>
                </w:tc>
            </w:tr>`,
			emptyContext
		);

		const insertedRowAsProp = new Row(
			{
				insertion: { author: 'Luis', date: date, id: 1 },
			},
			new Cell({})
		);
		const insertedRowAsNode = Row.fromNode(insertedRowNode, emptyContext);

		it('Row node has expected insertion objects', () => {
			expect(insertedRowAsNode.props.insertion).toEqual(
				insertedRowAsProp.props.insertion
			);
		});

		it('serializes and deserialized correctly', async () => {
			const testTable = new Table({});
			expect(
				serialize(await insertedRowAsProp.toNode([testTable]))
			).toEqual(
				serialize(
					create(
						`<tr xmlns="${NamespaceUri.w}">
							<trPr>
								<ins xmlns:ns1="${
									NamespaceUri.w
								}" ns1:id="1" ns1:author="Luis" ns1:date="${date.toISOString()}"/>
							</trPr>
            			</tr>`
					)
				)
			);
		});
	});
});
