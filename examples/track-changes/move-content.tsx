// /** @jsx  Docx.jsx */
// import Docx, { MoveRangeStart, Paragraph, Section, Text } from '../../mod.ts';
// import { Move } from '../../src/components/Move.ts';
// import { MoveRangeEnd } from '../../src/components/MoveRangeEnd.ts';

// // Create a new .docx file with track changes enabled.
// const docxFile = Docx.fromNothing().withSettings({
// 	isTrackChangesEnabled: true,
// });

// // Next, we'll create some new instances of our various objects to create our content.
// const date = new Date();

// const paragraph1 = new Paragraph(
// 	{},
// 	new Text({}, 'This is my original, first paragraph.')
// );

// const paragraph2 = new Paragraph(
// 	{},
// 	new Text({}, 'This will be moved from being paragraph 2, to paragraph 1.')
// );

// const moveToStart = new MoveRangeStart({
// 	id: 0,
// 	type: 'to',
// 	author: 'Gabe',
// 	// The name of a move to should match that of a move from. Word will still open the document if the two
// 	// do not match, but it will appear as an extra addition and movement.
// 	name: 'Move_0',
// 	date: date,
// });

// // const moveTo = new Move(
// // 	{
// // 		id: 1,
// // 		type: 'to',
// // 		author: 'Gabe',
// // 		date: date,
// // 	},
// // 	paragraph2
// // );

// const moveToEnd = new MoveRangeEnd({
// 	id: 0,
// 	type: 'to',
// });

// const moveFromStart = new MoveRangeStart({
// 	id: 2,
// 	type: 'from',
// 	author: 'Gabe',
// 	name: 'Move_0',
// 	date: date,
// });

// const moveFrom = new Move(
// 	{
// 		id: 2,
// 		type: 'from',
// 		author: 'Gabe',
// 		date: date,
// 	}
// );

// const moveFromEnd = new MoveRangeEnd({
// 	id: 2,
// 	type: 'from',
// });

// // Create a section that contains our content moves.
// const section = new Section(
// 	{},
// 	moveToStart,
// 	moveTo,
// 	moveToEnd,
// 	paragraph1,
// 	moveFromStart,
// 	moveFrom,
// 	moveFromEnd
// );

// // Add that section to our existing Word document.
// docxFile.document.set(section);

// // And write it to a file.
// docxFile.toFile('move-content.docx');
