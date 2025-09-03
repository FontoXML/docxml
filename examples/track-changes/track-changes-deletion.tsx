/** @jsx  Docx.jsx */
import Docx, {
	Deletion,
	FootnoteReference,
	Paragraph,
	Section,
	Text,
	type FootnoteProps,
} from '../../mod.ts';

// Create a new .docx file with track changes enabled.
const docxFile = Docx.fromNothing().withSettings({
	isTrackChangesEnabled: true,
});

const date = new Date();

const footnoteProps: FootnoteProps = {
	numberingFormat: 'lowerRoman',
	position: 'pageBottom',
	restart: 'continuous',
};

const footnoteReferenceStyleName = 'FootnoteReference';

// Create a footnote reference that will be deleted.
const newFootnoteReference = await docxFile.document.footnotes.add(
	new Paragraph({}, new Text({}, 'Perhaps not so old.')),
	footnoteReferenceStyleName
);

// Create a new deleted paragraph
const testDeletedParagraph = new Paragraph(
	{ pilcrow: { deletion: { author: 'Ángel', date: date, id: 1 } } },
	new Deletion(
		{ author: 'Ángel', date: date, id: 1 },
		new Text({}, 'This is just a test.'),
		new FootnoteReference({
			id: newFootnoteReference,
			style: 'FootnoteReference',
		})
	)
);

// Create a section as the parent of our new paragraph.
const testSection = new Section(
	{ footnotes: footnoteProps },
	testDeletedParagraph
);

// Set that section as the content of our document.
docxFile.document.set(testSection);

// Save our document.
await docxFile.toFile('track-changes-deletion.docx');

// Alternatively, you can use JSX:
await Docx.fromJsx(
	<Section footnotes={footnoteProps}>
		<Paragraph
			pilcrow={{ deletion: { id: 1, author: 'ines', date: new Date() } }}
		>
			<Deletion id={0} author="Ines" date={new Date()}>
				<Text>my old</Text>
				<FootnoteReference
					id={newFootnoteReference}
					style={footnoteReferenceStyleName}
				/>
				<Text>friend.</Text>
			</Deletion>
		</Paragraph>
	</Section>
).toFile('track-changes-deletion-jsx.docx');
