import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Cell, Paragraph, Row, Table, Text } from '../../../../mod.ts'; // ← your public barrel
import { Archive } from '../../../classes/src/Archive.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { create, serialize } from '../../../utilities/src/dom.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';

const date = new Date();
const emptyContext: ComponentContext = {
	archive: new Archive(),
	relationships: null,
};

describe('Cell-level track-changes (<cellIns>/<cellDel>)', () => {
	const tableNode = create(
		`<w:tbl xmlns:w="${NamespaceUri.w}">
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:cellDel w:id="1" w:author="Carlos" w:date="${date.toISOString()}"/>
					</w:tcPr>
					<w:p><w:r><w:t xml:space="preserve">Bye1!</w:t></w:r></w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:cellIns w:id="2" w:author="Carlos" w:date="${date.toISOString()}"/>
					</w:tcPr>
					<w:p><w:r><w:t xml:space="preserve">Hello1!</w:t></w:r></w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:cellDel w:id="3"/>
					</w:tcPr>
					<w:p><w:r><w:t xml:space="preserve">Bye2!</w:t></w:r></w:p>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:cellIns w:id="4"/>
					</w:tcPr>
					<w:p><w:r><w:t xml:space="preserve">Hello2!</w:t></w:r></w:p>
				</w:tc>
			</w:tr>
		</w:tbl>`,
		emptyContext
	);

	const tableObject = new Table(
		{},
		new Row(
			{},
			new Cell(
				{ deletion: { id: 1, author: 'Carlos', date } },
				new Paragraph({}, new Text({}, 'Bye1!'))
			),
			new Cell(
				{ insertion: { id: 2, author: 'Carlos', date } },
				new Paragraph({}, new Text({}, 'Hello1!'))
			),
			new Cell(
				{ deletion: { id: 3 } },
				new Paragraph({}, new Text({}, 'Bye2!'))
			),
			new Cell(
				{ insertion: { id: 4 } },
				new Paragraph({}, new Text({}, 'Hello2!'))
			)
		)
	);

	const tableFromNode = Table.fromNode(tableNode, emptyContext)!;

	it('parses <cellDel>/<cellIns> into cell props', () => {
		const [firstRow] = tableFromNode.children;
		const [delCell1, insCell1, delCell2, insCell2] =
			firstRow.children as Cell[];

		expect(delCell1.props.deletion).toEqual({
			id: 1,
			author: 'Carlos',
			date,
		});
		expect(insCell1.props.insertion).toEqual({
			id: 2,
			author: 'Carlos',
			date,
		});
		expect(delCell2.props.deletion).toEqual({
			id: 3,
		});
		expect(insCell2.props.insertion).toEqual({
			id: 4,
		});
	});

	it('serialises and deserialises correctly', async () => {
		/* what the API generates */
		const generated = serialize(await tableObject.toNode([]));

		/* what we expect – note:
     • default namespace on <tbl>, <tr>, <tc>, …
     • a <tblPr/> skeleton (the renderer always emits one)
     • ns1 / ns2 prefixes only on the change markers */
		const expected = serialize(
			create(
				`<tbl xmlns="${NamespaceUri.w}">
					<tblPr/>
					<tr>
						<tc>
							<tcPr>
								<cellDel ns1:id="1" ns1:author="Carlos" ns1:date="${date.toISOString()}" xmlns:ns1="${
					NamespaceUri.w
				}"/>
							</tcPr>
							<p>
								<r>
									<t xml:space="preserve">Bye1!</t>
								</r>
							</p>
						</tc>
						<tc>
							<tcPr>
								<cellIns ns2:id="2" ns2:author="Carlos" ns2:date="${date.toISOString()}" xmlns:ns2="${
					NamespaceUri.w
				}"/>
							</tcPr>
							<p>
								<r>
									<t xml:space="preserve">Hello1!</t>
								</r>
							</p>
						</tc>
						<tc>
							<tcPr>
								<cellDel ns3:id="3" xmlns:ns3="${NamespaceUri.w}"/>
							</tcPr>
							<p>
								<r>
									<t xml:space="preserve">Bye2!</t>
								</r>
							</p>
						</tc>
						<tc>
							<tcPr>
								<cellIns ns4:id="4" xmlns:ns4="${NamespaceUri.w}"/>
							</tcPr>
							<p>
								<r>
									<t xml:space="preserve">Hello2!</t>
								</r>
							</p>
						</tc>
					</tr>
				</tbl>`
			)
		);

		expect(generated).toEqual(expected);
	});
});
