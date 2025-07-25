/** @jsx  Docx.jsx */
import Docx, { Paragraph, Section, Text } from '../../mod.ts';

// Create a new .docx file with track changes enabled.
const docxFile = Docx.fromNothing().withSettings({
	isTrackChangesEnabled: true,
});

const date = new Date();

// Create a new inserted paragraph
const testParagraph = new Paragraph(
	{ pilcrow: { insertion: { author: 'Luis', date: date, id: 1 } } },
	new Text({}, ' my old friend.')
);

// Create a section as the parent of our new paragraph.
const testSection = new Section({}, testParagraph);

// Set that section as the content of our document.
docxFile.document.set(testSection);

// Save our document.
await docxFile.toFile('track-changes-paragraph.docx');

// Alternatively, you can use JSX:
await Docx.fromJsx(
	<Paragraph
		pilcrow={{ insertion: { id: 1, author: 'ines', date: new Date() } }}
	>
		{' '}
		my old friend.
	</Paragraph>
).toFile('track-changes-paragraph-jsx.docx');
