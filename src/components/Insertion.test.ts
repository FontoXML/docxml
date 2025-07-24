import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';
import { Archive } from '../classes/Archive.ts';
import type { ComponentContext } from '../classes/Component.ts';
import { create, serialize } from '../utilities/dom.ts';
import { NamespaceUri } from '../utilities/namespaces.ts';
import { Cell } from './Cell.ts';
import { Row } from './Row.ts';
import { Table } from './Table.ts';

describe('Insertion', () => {
	const date = new Date();

	const emptyContext: ComponentContext = {
		archive: new Archive(),
		relationships: null,
	};

	describe('Inserted run content', () => {
		/** 			const moveToObject = new Row(
			{}
			new Insertion(
				{
					id: 0,
					date: date,
					author: 'Gabe',
					type: 'to',
				}
			)
		);
				const newRowInsertion = Row.fromNode(insertedRowNode, emptyContext);

		**/
	});
	describe('Inserted numbering properties', () => {});
	describe('Inserted paragraph', () => {});
	describe('Inserted table row', () => {
		const rowWithInsertionNode = create(
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

		const rowWithInsertionAsProp = new Row(
			{
				insertion: { author: 'Luis', date: date, id: 1 },
			},
			new Cell({})
		);
		const rowWithInsertionAsNode = Row.fromNode(
			rowWithInsertionNode,
			emptyContext
		);

		it('Row node has expected insertion objects', () => {
			expect(rowWithInsertionAsNode.props.insertion).toEqual(
				rowWithInsertionAsProp.props.insertion
			);
		});

		it('serializes and deserialized correctly', async () => {
			const testTable = new Table({});
			expect(
				serialize(await rowWithInsertionAsProp.toNode([testTable]))
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

		//toNode() rowWithInsertionAsProp y un create comparar
	});
});
