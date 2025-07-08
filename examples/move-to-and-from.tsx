/** @jsx  Docx.jsx */
import Docx, { MoveRangeStart, Paragraph, Section, Text } from '../mod.ts';
import { Move } from '../src/components/Move.ts';
import { MoveRangeEnd } from '../src/components/MoveRangeEnd.ts';

// Create a new .docx file with track changes enabled.
const docxFile = Docx.fromNothing().withSettings({
	isTrackChangesEnabled: true,
});

// Next, we'll create some new instances of our various objects to create our content.
const date = new Date();

const paragraph1 = new Paragraph(
	{},
	new Text({}, 'This is my original, first paragraph.')
);

const paragraph2 = new Paragraph(
	{},
	new Text({}, 'This will be moved from being paragraph 2, to paragraph 1.')
);

const moveToStart = new MoveRangeStart({
	id: 0,
	type: 'to',
	author: 'Gabe',
	name: 'Move_0',
	date: date,
});

const moveTo = new Move(
	{
		id: 1,
		type: 'to',
		author: 'Gabe',
		date: date,
	},
	paragraph2
);

const moveToEnd = new MoveRangeEnd({
	id: 0,
	type: 'to',
});

const moveFromStart = new MoveRangeStart({
	id: 2,
	type: 'from',
	author: 'Gabe',
	name: 'Move_0',
	date: date,
});

const moveFrom = new Move(
	{
		id: 2,
		type: 'to',
		author: 'Gabe',
		date: date,
	},
	paragraph2
);

const moveFromEnd = new MoveRangeEnd({
	id: 2,
	type: 'from',
});

// Create a section that contains our content moves.
const section = new Section(
	{},
	moveToStart,
	moveTo,
	moveToEnd,
	paragraph1,
	moveFromStart,
	moveFrom,
	moveFromEnd
);

// Add that section to our existing Word document.
docxFile.document.set(section);

// And write it to a file.
docxFile.toFile('move-content.docx');

// Alternatively, you can use JSX:
const jsxWordDoc = Docx.fromNothing().withSettings({
	isTrackChangesEnabled: true,
});

jsxWordDoc.document.set(
	<Section>
		<Paragraph>
			<MoveRangeStart
				id={2}
				author="Gabe"
				date={date}
				name="Move_1"
				type="to"
			/>
			<Move type="to" id={2} date={date} author="Gabe">
				<Text>This is some more moved text.</Text>
			</Move>
			<MoveRangeEnd id={2} type="to" />
			<Text>This is a bit of sample text.</Text>
			<MoveRangeStart
				id={3}
				author="Gabe"
				date={date}
				name="Move_1"
				type="from"
			/>
			<Move type="from" id={3} date={date} author="Gabe">
				<Text>This is some more moved text.</Text>
			</Move>
			<MoveRangeEnd id={3} type="from" />
		</Paragraph>
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

jsxWordDoc.toFile('move-content-jsx.docx');
