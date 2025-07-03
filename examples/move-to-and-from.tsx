/** @jsx  Docx.jsx */
import Docx, { MoveRangeStart, Paragraph, Section, Text } from '../mod.ts';
import { Move } from '../src/components/Move.ts';
import { MoveRangeEnd } from '../src/components/MoveRangeEnd.ts';

// Create a new .docx file with track changes enabled.
const docxFile = Docx.fromNothing().withSettings({
	isTrackChangesEnabled: true,
});

const date = new Date();

// Alternatively, you can use JSX:
docxFile.document.set(
	<Section>
		<MoveRangeStart
			id={1}
			type="to"
			date={date}
			author="Gabe"
			name="Move_0"
		/>
		<Move type="to" id={1} date={date} author="Gabe">
			<Paragraph>
				<Text>This is my moved text.</Text>
			</Paragraph>
		</Move>
		<MoveRangeEnd id={1} type="to" />
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
		<Move type="from" id={0} date={date} author="Gabe">
			<Paragraph>
				<Text>This is my moved text.</Text>
			</Paragraph>
		</Move>
		<MoveRangeEnd id={0} type="from" />
	</Section>
);

docxFile.toFile('move-content-jsx.docx');
