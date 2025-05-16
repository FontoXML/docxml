/** @jsx  Docx.jsx */
import Docx, {
	type FootnoteProps,
	Footnote,
	Paragraph,
	Section,
} from '../mod.ts';
import { Text } from '../src/components/Text.ts';
import { pt } from '../src/utilities/length.ts';

const docxFile = Docx.fromNothing();

const footnoteProps: FootnoteProps = {
	numberingFormat: 'lowerRoman',
	position: 'beneathText',
	restart: 'continuous',
	styleName: 'FootnoteText',
	referenceStyleName: 'FootnoteReference',
};

const footnote1 = docxFile.document.footnotes.add(
	new Paragraph({}, new Text({}, 'Hello, this is a footnote.')),
	'normal',
	footnoteProps.styleName,
	footnoteProps.referenceStyleName
);
const footnote2 = docxFile.document.footnotes.add(
	new Paragraph({}, new Text({}, 'And this is an additional footnote.')),
	'normal',
	footnoteProps.styleName,
	footnoteProps.referenceStyleName
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
			<Footnote
				id={footnote1.id}
				styleName={footnoteProps.styleName}
				referenceStyleName={footnoteProps.referenceStyleName}
			/>
			<Footnote
				id={footnote2.id}
				styleName={footnoteProps.styleName}
				referenceStyleName={footnoteProps.referenceStyleName}
			/>
		</Paragraph>
	</Section>
);

await docxFile.toFile('footnotes.docx');
