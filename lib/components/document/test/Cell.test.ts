import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Archive } from '../../../classes/src/Archive.ts';
import { Bookmarks } from '../../../classes/src/Bookmarks.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { create } from '../../../utilities/src/dom.ts';
import { twip } from '../../../utilities/src/length.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';
import {
	evaluateXPathToArray,
	evaluateXPathToFirstNode,
	evaluateXPathToNodes,
} from '../../../utilities/src/xquery.ts';
import { Cell } from '../src/Cell.ts';
import { Row } from '../src/Row.ts';
import { Table } from '../src/Table.ts';

const emptyContext: ComponentContext = {
	archive: new Archive(),
	relationships: null,
	bookmarks: new Bookmarks(),
};

describe('Cell', () => {
	const dom = create(`<w:tbl xmlns:w="${NamespaceUri.w}">
		<w:tblGrid>
			<w:gridCol w:w="4319" />
			<w:gridCol w:w="4319" />
		</w:tblGrid>
		<w:tr>
			<w:trPr />
			<w:tc xid="cell-1">
				<w:p>
					<w:r>
						<w:t xml:space="preserve">A 1</w:t>
					</w:r>
				</w:p>
			</w:tc>
			<w:tc xid="cell-2">
				<w:p>
					<w:r>
						<w:t xml:space="preserve">B 1</w:t>
					</w:r>
				</w:p>
			</w:tc>
		</w:tr>
		<w:tr>
			<w:trPr />
			<w:tc xid="cell-3">
				<w:tcPr>
					<w:gridSpan w:val="2" />
				</w:tcPr>
				<w:p>
					<w:r>
						<w:t xml:space="preserve">A+B 2</w:t>
					</w:r>
				</w:p>
			</w:tc>
		</w:tr>
	</w:tbl>`);

	describe('Cell 1', () => {
		const cell = Cell.fromNode(
			// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
			evaluateXPathToFirstNode('.//*[@xid="cell-1"]', dom)!,
			emptyContext
		);
		it('Colspan', () => expect(cell?.props.colSpan).toBe(1));
		it('Rowspan', () => expect(cell?.props.rowSpan).toBe(1));
	});
	describe('Cell 2', () => {
		const cell = Cell.fromNode(
			// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
			evaluateXPathToFirstNode('.//*[@xid="cell-2"]', dom)!,
			emptyContext
		);
		it('Colspan', () => expect(cell?.props.colSpan).toBe(1));
		it('Rowspan', () => expect(cell?.props.rowSpan).toBe(1));
	});
	describe('Cell 3', () => {
		const cell = Cell.fromNode(
			// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
			evaluateXPathToFirstNode('.//*[@xid="cell-3"]', dom)!,
			emptyContext
		);
		it('Colspan', () => expect(cell?.props.colSpan).toBe(2));
		it('Rowspan', () => expect(cell?.props.rowSpan).toBe(1));
	});
});

describe('Cell - tcMar (per-cell margins)', () => {
	const dom = create(`<w:tbl xmlns:w="${NamespaceUri.w}">
		<w:tr>
			<w:tc xid="modern">
				<w:tcPr>
					<w:tcMar>
						<w:top w:w="28" w:type="dxa"/>
						<w:start w:w="14" w:type="dxa"/>
						<w:bottom w:w="28" w:type="dxa"/>
						<w:end w:w="14" w:type="dxa"/>
					</w:tcMar>
				</w:tcPr>
				<w:p/>
			</w:tc>
			<w:tc xid="legacy">
				<w:tcPr>
					<w:tcMar>
						<w:left w:w="14" w:type="dxa"/>
						<w:right w:w="14" w:type="dxa"/>
					</w:tcMar>
				</w:tcPr>
				<w:p/>
			</w:tc>
		</w:tr>
	</w:tbl>`);

	describe('modern start/end', () => {
		const cell = Cell.fromNode(
			// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
			evaluateXPathToFirstNode('.//*[@xid="modern"]', dom)!,
			emptyContext
		);
		it('parses all four sides', () => {
			expect(cell?.props.margin?.top?.twip).toBe(28);
			expect(cell?.props.margin?.start?.twip).toBe(14);
			expect(cell?.props.margin?.bottom?.twip).toBe(28);
			expect(cell?.props.margin?.end?.twip).toBe(14);
		});
	});

	describe('legacy left/right', () => {
		const cell = Cell.fromNode(
			// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
			evaluateXPathToFirstNode('.//*[@xid="legacy"]', dom)!,
			emptyContext
		);
		it('maps left/right onto start/end', () => {
			expect(cell?.props.margin?.start?.twip).toBe(14);
			expect(cell?.props.margin?.end?.twip).toBe(14);
		});
	});
});

describe('Cell - with colspan', () => {
	const tableNode = create(`
		<w:tbl xmlns:w="${NamespaceUri.w}">
			<w:tblGrid>
				<w:gridCol w:w="1129"/>
				<w:gridCol w:w="2268"/>
				<w:gridCol w:w="4536"/>
				<w:gridCol w:w="1083"/>
			</w:tblGrid>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="1129" w:type="dxa"/>
					</w:tcPr>
					<w:p/>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="7887" w:type="dxa"/>
						<w:gridSpan w:val="3"/>
					</w:tcPr>
				<w:p/>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="1129" w:type="dxa"/>
					</w:tcPr>
					<w:p/>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="7887" w:type="dxa"/>
						<w:gridSpan w:val="3"/>
					</w:tcPr>
				<w:p/>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:trPr>
				<w:trHeight w:val="949"/>
				</w:trPr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="1129" w:type="dxa"/>
					</w:tcPr>
					<w:p/>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="7887" w:type="dxa"/>
						<w:gridSpan w:val="3"/>
					</w:tcPr>
				<w:p/>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="1129" w:type="dxa"/>
					</w:tcPr>
					<w:p/>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="2268" w:type="dxa"/>
					</w:tcPr>
					<w:p/>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="4536" w:type="dxa"/>
					</w:tcPr>
					<w:p/>
				</w:tc>
				<w:tc>
					<w:tcPr>
						<w:tcW w:w="1083" w:type="dxa"/>
					</w:tcPr>
					<w:p/>
				</w:tc>
			</w:tr>
		</w:tbl>
	`);

	it('cell width', async () => {
		const fromNode = Table.fromNode(tableNode, emptyContext);
		const toNode = await fromNode.toNode([]);

		const originalCellWidths = evaluateXPathToNodes(
			`./descendant::Q{${NamespaceUri.w}}tcW`,
			tableNode
		).map((tcW) => (tcW as Element).getAttributeNS(NamespaceUri.w, 'w'));
		const newCellWidths = evaluateXPathToNodes(
			`./descendant::Q{${NamespaceUri.w}}tcW`,
			toNode
		).map((tcW) => (tcW as Element).getAttributeNS(NamespaceUri.w, 'w'));

		expect(originalCellWidths).toEqual(newCellWidths);
	});
});

describe('Cell - with borders', () => {
	const tableNode = create(`
		<w:tbl xmlns:w="${NamespaceUri.w}">
			<w:tblGrid>
				<w:gridCol w:w="1129"/>
			</w:tblGrid>
			<w:tr>
				<w:tc>
					<w:tcPr>
						<w:tcBorders>
							<w:top w:val="double" w:sz="24" w:space="0" w:color="FF0000"/>
						</w:tcBorders>
					</w:tcPr>
					<w:p/>
				</w:tc>
			</w:tr>
		</w:tbl>
	`);

	it('cell width', async () => {
		const fromNode = Table.fromNode(tableNode, emptyContext);
		const toNode = await fromNode.toNode([]);

		const originalBorder = evaluateXPathToFirstNode<Element>(
			`./descendant::Q{${NamespaceUri.w}}tcBorders`,
			tableNode
		);
		const newBorder = evaluateXPathToFirstNode<Element>(
			`./descendant::Q{${NamespaceUri.w}}tcBorders`,
			toNode
		);

		expect(originalBorder?.getAttributeNS(NamespaceUri.w, 'val')).toEqual(
			newBorder?.getAttributeNS(NamespaceUri.w, 'val')
		);
		expect(originalBorder?.getAttributeNS(NamespaceUri.w, 'sz')).toEqual(
			newBorder?.getAttributeNS(NamespaceUri.w, 'sz')
		);
		expect(originalBorder?.getAttributeNS(NamespaceUri.w, 'space')).toEqual(
			newBorder?.getAttributeNS(NamespaceUri.w, 'space')
		);
		expect(originalBorder?.getAttributeNS(NamespaceUri.w, 'color')).toEqual(
			newBorder?.getAttributeNS(NamespaceUri.w, 'color')
		);
	});
});

describe('Cell margins of vertically merged rows', () => {
	// Mirrors a Word template where the merged cell has different margins
	// in the rows it spans.
	const dom = create(`<w:tbl xmlns:w="${NamespaceUri.w}">
		<w:tblGrid>
			<w:gridCol w:w="2879" />
			<w:gridCol w:w="7160" />
		</w:tblGrid>
		<w:tr>
			<w:tc xid="merged">
				<w:tcPr>
					<w:vMerge w:val="restart"/>
					<w:tcMar><w:top w:w="43" w:type="dxa"/><w:bottom w:w="43" w:type="dxa"/></w:tcMar>
				</w:tcPr>
				<w:p/>
			</w:tc>
			<w:tc><w:p/></w:tc>
		</w:tr>
		<w:tr>
			<w:tc>
				<w:tcPr>
					<w:vMerge/>
					<w:tcMar><w:top w:w="14" w:type="dxa"/><w:bottom w:w="14" w:type="dxa"/></w:tcMar>
				</w:tcPr>
				<w:p/>
			</w:tc>
			<w:tc><w:p/></w:tc>
		</w:tr>
		<w:tr>
			<w:tc>
				<w:tcPr>
					<w:vMerge/>
				</w:tcPr>
				<w:p/>
			</w:tc>
			<w:tc><w:p/></w:tc>
		</w:tr>
	</w:tbl>`);

	it('reads the margins of each spanned row', () => {
		const cell = Cell.fromNode(
			// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
			evaluateXPathToFirstNode('.//*[@xid="merged"]', dom)!,
			emptyContext
		);

		expect(cell?.props.rowSpan).toBe(3);
		expect(cell?.props.margin?.top?.twip).toBe(43);
		expect(cell?.props.spannedRowMargins).toHaveLength(2);
		expect(cell?.props.spannedRowMargins?.[0]?.top?.twip).toBe(14);
		expect(cell?.props.spannedRowMargins?.[0]?.bottom?.twip).toBe(14);
		expect(cell?.props.spannedRowMargins?.[1]).toBeNull();
	});

	it('writes the margins of each spanned row', async () => {
		const table = Table.fromNode(dom, emptyContext);
		const node = await table.toNode([]);

		expect(
			evaluateXPathToArray(
				`array { ./*[local-name() = "tr"]/*[local-name() = "tc"][1]/string(
					./*[local-name() = "tcPr"]/*[local-name() = "tcMar"]/*[local-name() = "top"]/@*[local-name() = "w"]
				) }`,
				node
			)
		).toEqual(['43', '14', '']);
	});

	it('uses the first row margins when spannedRowMargins is not set', async () => {
		const table = new Table(
			{ columnWidths: [twip(2879), twip(7160)] },
			new Row(
				{},
				new Cell({
					rowSpan: 2,
					margin: { top: twip(43), bottom: twip(43) },
				}),
				new Cell({})
			),
			new Row({}, new Cell({}))
		);

		const node = await table.toNode([]);

		expect(
			evaluateXPathToArray(
				`array { ./*[local-name() = "tr"]/*[local-name() = "tc"][1]/string(
					./*[local-name() = "tcPr"]/*[local-name() = "tcMar"]/*[local-name() = "top"]/@*[local-name() = "w"]
				) }`,
				node
			)
		).toEqual(['43', '43']);
	});

	it('does not set spannedRowMargins when no spanned row has its own margins', () => {
		const noMarginsDom = create(`<w:tbl xmlns:w="${NamespaceUri.w}">
			<w:tblGrid>
				<w:gridCol w:w="2879" />
			</w:tblGrid>
			<w:tr>
				<w:tc xid="merged">
					<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
					<w:p/>
				</w:tc>
			</w:tr>
			<w:tr>
				<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
			</w:tr>
		</w:tbl>`);

		const cell = Cell.fromNode(
			// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
			evaluateXPathToFirstNode('.//*[@xid="merged"]', noMarginsDom)!,
			emptyContext
		);

		expect(cell?.props.rowSpan).toBe(2);
		expect(cell && 'spannedRowMargins' in cell.props).toBe(false);
	});
});
