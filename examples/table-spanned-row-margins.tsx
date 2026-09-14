// deno-lint-ignore-file jsx-key
/** @jsx Docx.jsx */

import Docx, {
	type CellMargin,
	Cell,
	Paragraph,
	pt,
	Row,
	Section,
	Table,
	twip,
} from '../mod.ts';

const FIRST_ROW_MARGIN: CellMargin = { top: pt(20), bottom: pt(20) };
const SPANNED_ROW_MARGIN: CellMargin = { top: pt(0), bottom: pt(0) };
const BORDER = { type: 'single', width: pt(0.5), color: '000000' } as const;

/**
 * A header-like table with the first column merged over three rows. The first row has larger top
 * and bottom margins, while the spanned rows can have smaller margins through `spannedRowMargins`.
 */
function createTable(spannedRowMargins?: Array<CellMargin | null>) {
	return (
		<Table
			columnWidths={[twip(2879), twip(7160)]}
			cellPadding={{ start: twip(14), end: twip(14) }}
			borders={{
				top: BORDER,
				start: BORDER,
				bottom: BORDER,
				end: BORDER,
				insideH: BORDER,
				insideV: BORDER,
			}}
		>
			<Row>
				<Cell
					rowSpan={3}
					verticalAlignment='center'
					shading={{ background: 'D9D9D9' }}
					margin={FIRST_ROW_MARGIN}
					spannedRowMargins={spannedRowMargins}
				>
					<Paragraph>Logo</Paragraph>
				</Cell>
				<Cell verticalAlignment='center' margin={FIRST_ROW_MARGIN}>
					<Paragraph alignment='right'>Company</Paragraph>
				</Cell>
			</Row>
			<Row>
				<Cell verticalAlignment='center'>
					<Paragraph>
						Document Title: &lt;Document Title&gt;
					</Paragraph>
				</Cell>
			</Row>
			<Row>
				<Cell verticalAlignment='center'>
					<Paragraph>Doc.No.: &lt;Doc.No.&gt;</Paragraph>
				</Cell>
			</Row>
		</Table>
	);
}

await Docx.fromJsx([
	<Section>
		<Paragraph>
			Without spannedRowMargins: rows 2 and 3 repeat the 20pt top and
			bottom margins of the merged grey cell, so they are much taller than
			their text.
		</Paragraph>
		{createTable()}
		<Paragraph />
		<Paragraph>
			With spannedRowMargins: rows 2 and 3 use their own 0pt margins, so
			they are only as tall as their text.
		</Paragraph>
		{createTable([SPANNED_ROW_MARGIN, SPANNED_ROW_MARGIN])}
	</Section>,
]).toFile('table-spanned-row-margins.docx');
