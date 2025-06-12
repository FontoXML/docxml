/** @jsx  Docx.jsx */
import Docx, {
	Paragraph,
	Section,
	Text,
	TextAddition,
	TextDeletion,
} from '../mod.ts';

const docxFile = Docx.fromNothing().withSettings({
	isTrackChangesEnabled: true,
});
const testParagraph = new Paragraph(
	{},
	new Text({}, 'Hello, '),
	new TextDeletion(
		{
			id: 2,
			author: 'Gabe',
			date: new Date(),
		},
		new Text({}, 'nighttime')
	),
	new Text({}, " my old friend."),
	new TextAddition(
		{ id: 1, author: 'Gabe', date: new Date() },
		new Text({}, "I've come to talk with you again.")
	),
);

const testSection = new Section({}, testParagraph);

docxFile.document.set(testSection);
await docxFile.toFile('track-changes.docx');
