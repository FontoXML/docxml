/** @jsx  Docx.jsx */
import Docx, { Cell, Image, Paragraph, Row, Section, Table } from '../mod.ts';
import {
	type FootnoteProps,
	FootnoteReference,
} from '../src/components/Footnote.ts';
import { Text } from '../src/components/Text.ts';
import { RelationshipType } from '../src/enums.ts';
import { inch, pt } from '../src/utilities/length.ts';

const docxFile = Docx.fromNothing();

const footnoteProps: FootnoteProps = {
	numberingFormat: 'lowerRoman',
	position: 'beneathText',
	restart: 'continuous',
	styleName: 'FootnoteText',
	referenceStyleName: 'FootnoteReference',
};

const newImage = new Image({
	data: Deno.readFile('test/spacekees.jpeg'),
	width: inch(1),
	height: inch(1),
});

await newImage.ensureRelationship(docxFile.document.relationships);

const testCell = new Cell(
	{},
	new Paragraph(
		{},
		new Text({ style: 'FootnoteText' }, 'This is a table cell')
	)
);

const testRow = new Row({}, testCell);

const testTable = new Table({}, testRow);

const footnote1 = docxFile.document.footnotes.add(
	new Paragraph({}, new Text({}, 'Hello, this is a footnote.')),
	'normal',
	footnoteProps.styleName,
	footnoteProps.referenceStyleName
);
const footnote2 = docxFile.document.footnotes.add(
	[
		new Paragraph({}, new Text({}, 'And this is an additional footnote.')),
		new Paragraph(
			{},
			new Text({}, 'And it will have more than one paragraph')
		),
	],
	'normal',
	footnoteProps.styleName,
	footnoteProps.referenceStyleName
);

const footnote3 = docxFile.document.footnotes.add(
	newImage,
	'normal',
	footnoteProps.styleName,
	footnoteProps.referenceStyleName
);

const footnote4 = docxFile.document.footnotes.add(
	testTable,
	'normal',
	'FootnoteText',
	'FootnoteReference'
);

docxFile.document.styles.add({
	id: 'FootnoteReference',
	name: 'FootnoteReference',
	type: 'character',
	text: {
		verticalAlign: 'superscript',
		fontSize: pt(10),
	},
});

docxFile.document.styles.add({
	id: 'FootnoteText',
	name: 'FootnoteText',
	type: 'paragraph',
	text: {
		fontSize: pt(9),
	},
});

docxFile.document.settings.set('footnoteProperties', {
	numberingFormat: 'lowerRoman',
	position: 'beneathText',
	restart: 'eachPage',
});

docxFile.document.set(
	<Section footnotes={footnoteProps}>
		<Paragraph>
			This is my first paragraph of text.
			<FootnoteReference
				id={footnote1.id}
				styleName={footnoteProps.styleName}
				referenceStyleName={footnoteProps.referenceStyleName}
			/>
			<FootnoteReference
				id={footnote2.id}
				styleName={footnoteProps.styleName}
				referenceStyleName={footnoteProps.referenceStyleName}
			/>
			<FootnoteReference
				id={footnote3.id}
				styleName={footnoteProps.styleName}
				referenceStyleName={footnoteProps.referenceStyleName}
			/>
			<FootnoteReference
				id={footnote4.id}
				styleName={footnoteProps.styleName}
				referenceStyleName={footnoteProps.referenceStyleName}
			/>
		</Paragraph>
	</Section>
);

await docxFile.toFile('footnotes.docx');
