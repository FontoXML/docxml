import Docx, {
	Move,
	MoveRangeEnd,
	MoveRangeStart,
	Paragraph,
	Section,
	Text,
} from '../../mod.ts';

// Create a new Word document with track changes enabled.
const docxFile = Docx.fromNothing().withSettings({
	isTrackChangesEnabled: true,
});

const date = new Date();

// Create an instance of a paragraph where the entire paragraph has been moved. This type of move
// has no range associated with it. It is assumed the parent paragraph is the extent of the range.
const moveToParagraph = new Paragraph(
	{
		pilcrow: {
			move: {
				id: 1,
				date: date,
				type: 'to',
				author: 'Gabe',
			},
		},
	},
	new MoveRangeStart({
		type: 'to',
		name: 'move_0',
		author: 'Gabe',
		date: date,
		id: 3,
	}),
	new Move(
		{
			type: 'to',
			id: 4,
			date: date,
			author: 'Gabe',
		},
		new Text({}, 'This is an example of some moved text.')
	),
	new MoveRangeEnd({
		type: 'to',
		id: 4,
	}),
	// Without date
	new MoveRangeStart({
		type: 'to',
		name: 'move_1',
		author: 'Inés',
		id: 5,
	}),
	new Move(
		{
			type: 'to',
			id: 5,
			author: 'Inés',
		},
		new Text({}, 'To the people who look at the stars and wish.')
	),
	new MoveRangeEnd({
		type: 'to',
		id: 5,
	}),
	// Without author
	new MoveRangeStart({
		type: 'to',
		name: 'move_2',
		date: date,
		id: 6,
	}),
	new Move(
		{
			type: 'to',
			id: 6,
			date: date,
		},
		new Text({}, 'There are different kinds of darkness.')
	),
	new MoveRangeEnd({
		type: 'to',
		id: 6,
	}),
	// Without author and date
	new MoveRangeStart({
		type: 'to',
		name: 'move_3',
		id: 7,
	}),
	new Move(
		{
			type: 'to',
			id: 7,
		},
		new Text(
			{},
			'No one was my master— but I might be master of everything, if I wished. If I dared.'
		)
	),
	new MoveRangeEnd({
		type: 'to',
		id: 7,
	})
);

const betweenParagraph = new Paragraph(
	{},
	new Text(
		{},
		'This will go before a completely move paragraph, but will show up as after it. '
	)
);

const moveFromParagraph = new Paragraph(
	{
		pilcrow: {
			move: {
				id: 1,
				date: date,
				type: 'from',
				author: 'Gabe',
			},
		},
	},
	new MoveRangeStart({
		type: 'from',
		name: 'move_0',
		author: 'Gabe',
		date: date,
		id: 3,
	}),
	new Move(
		{
			type: 'from',
			id: 4,
			date: date,
			author: 'Gabe',
		},
		new Text({}, 'This is an example of some moved text.')
	),
	new MoveRangeEnd({
		type: 'from',
		id: 4,
	}),
	// Without date
	new MoveRangeStart({
		type: 'from',
		name: 'move_1',
		author: 'Inés',
		id: 5,
	}),
	new Move(
		{
			type: 'from',
			id: 5,
			author: 'Inés',
		},
		new Text({}, 'To the people who look at the stars and wish.')
	),
	new MoveRangeEnd({
		type: 'from',
		id: 5,
	}),
	// Without author
	new MoveRangeStart({
		type: 'from',
		name: 'move_2',
		id: 6,
	}),
	new Move(
		{
			type: 'from',
			id: 6,
			date: date,
		},
		new Text({}, 'There are different kinds of darkness.')
	),
	new MoveRangeEnd({
		type: 'from',
		id: 6,
	}),
	// Without author and date
	new MoveRangeStart({
		type: 'from',
		name: 'move_3',
		id: 7,
	}),
	new Move(
		{
			type: 'from',
			id: 7,
		},
		new Text(
			{},
			'No one was my master— but I might be master of everything, if I wished. If I dared.'
		)
	),
	new MoveRangeEnd({
		type: 'from',
		id: 7,
	})
);

// Create an instance where text within a paragraph has been been moved. We will need to create a MoveRangeStart
// and MoveRangeEnd and Move for each piece of text that is moved.
const moveTextParagraph = new Paragraph(
	{},
	new MoveRangeStart({
		type: 'to',
		name: 'move_0',
		author: 'Gabe',
		date: date,
		id: 2,
	}),
	new Move(
		{
			type: 'to',
			author: 'Gabe',
			date: date,
			id: 2,
		},
		new Text({}, 'And this is some text that will move to be first.')
	),
	new MoveRangeEnd({
		type: 'to',
		id: 2,
	}),
	new Text({}, ' It will come from the middle of the text. '),
	new MoveRangeStart({
		type: 'from',
		name: 'move_0',
		author: 'Gabe',
		date: date,
		id: 3,
	}),
	new Move(
		{
			type: 'from',
			author: 'Gabe',
			date: date,
			id: 3,
		},
		new Text({}, 'And this is some text that will move to be first.')
	),
	new MoveRangeEnd({
		type: 'from',
		id: 3,
	})
);

// Add all three of the paragraphs we created to our document.
const section = new Section(
	{},
	moveToParagraph,
	betweenParagraph,
	moveFromParagraph,
	moveTextParagraph
);

docxFile.document.set(section);

// Write to a file.
docxFile.toFile('track-changes-move.docx');
