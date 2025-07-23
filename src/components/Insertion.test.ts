import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';
import { Archive } from '../classes/Archive.ts';
import type { ComponentContext } from '../classes/Component.ts';
import { create } from '../utilities/dom.ts';
import { NamespaceUri } from '../utilities/namespaces.ts';
import { Cell } from './Cell.ts';
import { Row } from './Row.ts';

describe('Insertion', () => {
	const date = new Date();

	const emptyContext: ComponentContext = {
		archive: new Archive(),
		relationships: null,
	};

	describe('Inserted run content', () => {});
	describe('Inserted numbering properties', () => {});
	describe('Inserted paragraph', () => {});
	describe('Inserted table row', () => {
		const insertedRowToAsPropNode = create(
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
		const insertedRowAsPropObject = new Row(
			{
				insertion: { author: 'Luis', date: date, id: 1 },
			},
			new Cell({})
		);
		const newInsertedRowToAsProp = Row.fromNode(
			insertedRowToAsPropNode,
			emptyContext
		);

		it('turns node into expected Row insertion objects', () => {
			expect(newInsertedRowToAsProp.props.insertion).toEqual(
				insertedRowAsPropObject.props.insertion
			);
		});
	});
});
