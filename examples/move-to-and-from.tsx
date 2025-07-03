/** @jsx  Docx.jsx */
import Docx, { MoveRangeStart, Paragraph, Section, Text } from '../mod.ts';

// Create a new .docx file with track changes enabled.
const docxFile = Docx.fromNothing().withSettings({
	isTrackChangesEnabled: true,
});

const date = new Date();

// Alternatively, you can use JSX:
docxFile.document.set(
	<Section>
		<Paragraph>
			<Text>This is my text.</Text>
		</Paragraph>
		<MoveRangeStart
			type="from"
			id={0}
			author="Gabe"
			name="Move_0"
			date={date}
		/>
	</Section>
);

docxFile.toFile('move-content-jsx.docx');
