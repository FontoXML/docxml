/** @jsx  Docx.jsx */
import Docx, { MoveRangeStart, Paragraph, Section, Text } from '../../mod.ts';
import { Move } from '../../src/components/Move.ts';
import { MoveRangeEnd } from '../../src/components/MoveRangeEnd.ts';

/**  We can use JSX to  create content moves as part of Word's track changes functionality.
 * To do so, we must wrap a set of <Move/> tags in between <MoveRangeStart /> and
 * <MoveRangeEnd /> tags.
 */
const jsxWordDoc = Docx.fromNothing().withSettings({
	isTrackChangesEnabled: true,
});

const date = new Date();

jsxWordDoc.document.set(
	<Section>
		<Paragraph>
			<MoveRangeStart
				type="to"
				id={4}
				name="Text_run_move"
				author="Gabe"
				date={date}
			/>
			<Move type="to" id={5} date={date} author="Gabe">
				<Text>This is moved text.</Text>
			</Move>
			<MoveRangeEnd id={4} type="to" />
			<Text> This is text between two moved texsts. </Text>
			<MoveRangeStart
				type="from"
				id={7}
				name="Text_run_move"
				author="Gabe"
				date={date}
			/>
			<Move type="from" id={5} date={date} author="Gabe">
				<Text>This is moved text.</Text>
			</Move>
			<MoveRangeEnd id={7} type="from" />
		</Paragraph>
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
