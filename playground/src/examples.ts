export type Example = { id: string; label: string; source: string };

export const examples: Example[] = [
	{
		id: 'hello',
		label: 'Hello World',
		source: `import Docx, { Paragraph, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.set(
		new Paragraph({}, new Text({}, 'Hello from the docxml playground.'))
	);

	return docx;
}
`,
	},
	{
		id: 'styled-text',
		label: 'Styled Text',
		source: `import Docx, { Paragraph, Text, pt } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.set(
		new Paragraph({ alignment: 'center' },
			new Text({ isBold: true, fontSize: pt(28) }, 'Bold & Centered'),
			new Text({}, ' \\u2014 '),
			new Text({ isItalic: true, color: '0066CC' }, 'Italic & Blue')
		)
	);

	return docx;
}
`,
	},
	{
		id: 'multiple-paragraphs',
		label: 'Multiple Paragraphs',
		source: `import Docx, { Paragraph, Text, pt } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.set([
		new Paragraph({}, new Text({ isBold: true, fontSize: pt(24) }, 'Document Title')),
		new Paragraph({}, new Text({}, 'First paragraph of content.')),
		new Paragraph({}, new Text({}, 'Second paragraph of content.')),
		new Paragraph({},
			new Text({}, 'Mixed: '),
			new Text({ isBold: true }, 'bold'),
			new Text({}, ', '),
			new Text({ isItalic: true }, 'italic'),
			new Text({}, ', and '),
			new Text({ isBold: true, isItalic: true }, 'both'),
			new Text({}, '.')
		),
	]);

	return docx;
}
`,
	},
	{
		id: 'table',
		label: 'Table',
		source: `import Docx, { Table, Row, Cell, Paragraph, Text, cm, pt } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.set(
		new Table(
			{
				columnWidths: [cm(3), cm(2.5), cm(2.5), cm(2.5), cm(2.5)],
				cellPadding: { top: pt(6), bottom: pt(6), start: pt(6), end: pt(6) },
				borders: {
					top: { type: 'single', width: pt(1), color: '666666' },
					bottom: { type: 'single', width: pt(1), color: '666666' },
					start: { type: 'single', width: pt(1), color: '666666' },
					end: { type: 'single', width: pt(1), color: '666666' },
					insideH: { type: 'dashed', width: pt(1), color: 'CCCCCC' },
					insideV: { type: 'dashed', width: pt(1), color: 'CCCCCC' },
				},
			},
			new Row({ isHeaderRow: true },
				new Cell({}, new Paragraph({}, new Text({ isBold: true }, 'Year'))),
				new Cell({}, new Paragraph({}, new Text({ isBold: true }, 'Q1'))),
				new Cell({}, new Paragraph({}, new Text({ isBold: true }, 'Q2'))),
				new Cell({}, new Paragraph({}, new Text({ isBold: true }, 'Q3'))),
				new Cell({}, new Paragraph({}, new Text({ isBold: true }, 'Q4')))
			),
			new Row({},
				new Cell({}, new Paragraph({}, new Text({}, '2023'))),
				new Cell({}, new Paragraph({}, new Text({}, '120'))),
				new Cell({}, new Paragraph({}, new Text({}, '145'))),
				new Cell({}, new Paragraph({}, new Text({}, '132'))),
				new Cell({}, new Paragraph({}, new Text({}, '158')))
			),
			new Row({},
				new Cell({}, new Paragraph({}, new Text({}, '2024'))),
				new Cell({}, new Paragraph({}, new Text({}, '165'))),
				new Cell({}, new Paragraph({}, new Text({}, '170'))),
				new Cell({}, new Paragraph({}, new Text({}, '155'))),
				new Cell({}, new Paragraph({}, new Text({}, '190')))
			)
		)
	);

	return docx;
}
`,
	},
	{
		id: 'sections',
		label: 'Sections & Page Sizes',
		source: `import Docx, { Section, Paragraph, Text, cm } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.set([
		new Section({ pageWidth: cm(21), pageHeight: cm(29.7) },
			new Paragraph({}, new Text({}, 'This is an A4 portrait page.'))
		),
		new Section({ pageWidth: cm(29.7), pageHeight: cm(21), pageOrientation: 'landscape' },
			new Paragraph({}, new Text({}, 'This is an A4 landscape page.'))
		),
		new Section({ pageWidth: cm(20), pageHeight: cm(20) },
			new Paragraph({}, new Text({}, 'This is a square page.'))
		),
	]);

	return docx;
}
`,
	},
	{
		id: 'hyperlinks',
		label: 'Hyperlinks & Bookmarks',
		source: `import Docx, { BookmarkRangeStart, BookmarkRangeEnd, Hyperlink, Paragraph, Section, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	const bookmark = docx.bookmarks.create();

	docx.document.set([
		new Section({},
			new Paragraph({},
				new Hyperlink({ bookmark },
					new Text({ color: '0563C1', isUnderlined: true }, 'Jump to bookmarked section')
				)
			)
		),
		new Section({},
			new BookmarkRangeStart({ bookmark }),
			new Paragraph({},
				new Hyperlink({ url: 'https://github.com/fontoxml/docxml' },
					new Text({ color: '0563C1', isUnderlined: true }, 'Visit docxml on GitHub')
				)
			),
			new BookmarkRangeEnd({ bookmark })
		),
	]);

	return docx;
}
`,
	},
	{
		id: 'comments',
		label: 'Comments',
		source: `import Docx, { Comment, CommentRangeStart, CommentRangeEnd, Paragraph, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.styles.add({
		id: 'CommentReference',
		type: 'character',
		paragraph: {},
	});

	const comment = docx.document.comments.add(
		{ author: 'Reviewer', date: new Date(), initials: 'R' },
		[new Paragraph({}, new Text({}, 'This needs revision.'))]
	);

	const reply = docx.document.comments.add(
		{ author: 'Author', date: new Date(), parentId: comment },
		[new Paragraph({}, new Text({}, 'I agree, will fix.'))]
	);

	docx.document.set(
		new Paragraph({},
			new Text({}, 'NSYNC is the '),
			new CommentRangeStart({ id: comment }),
			new CommentRangeStart({ id: reply }),
			new Text({}, 'greatest'),
			new Comment({ id: comment }),
			new Comment({ id: reply }),
			new CommentRangeEnd({ id: comment }),
			new CommentRangeEnd({ id: reply }),
			new Text({}, ' band in history.')
		)
	);

	return docx;
}
`,
	},
	{
		id: 'fields',
		label: 'Fields (Hyperlink)',
		source: `import Docx, { FieldDefinition, FieldNames, FieldRangeEnd, FieldRangeInstruction, FieldRangeSeparator, FieldRangeStart, Paragraph, Section, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.set(
		new Section({},
			new Paragraph({},
				new FieldRangeStart({}),
				new Text({},
					new FieldRangeInstruction({},
						new FieldDefinition({ name: FieldNames.HYPERLINK, value: 'http://www.google.com' })
					),
					new FieldRangeSeparator({}),
					'Click here to visit Google'
				),
				new FieldRangeEnd({})
			)
		)
	);

	return docx;
}
`,
	},
	{
		id: 'headers-footers',
		label: 'Headers & Footers',
		source: `import Docx, { Paragraph, Section, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	const header = docx.document.headers.add('word/header1.xml', [
		new Paragraph({}, new Text({ isBold: true }, 'Document Header')),
	]);

	const footer = docx.document.footers.add(
		'word/footer1.xml',
		new Paragraph({}, new Text({ color: '888888' }, 'Page footer — docxml playground'))
	);

	docx.document.set(
		new Section({ headers: header, footers: footer },
			new Paragraph({}, new Text({}, 'This page has a header and a footer.')),
			new Paragraph({}, new Text({}, 'Check the top and bottom of the page in Word.'))
		)
	);

	return docx;
}
`,
	},
	{
		id: 'protected',
		label: 'Protected Document',
		source: `import Docx, { Paragraph, Section, Text, cm } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.settings.set('documentProtection', {
		edit: 'readOnly',
		enforcement: true,
	});

	docx.document.set(
		new Section({ pageWidth: cm(21), pageHeight: cm(29.7) },
			new Paragraph({}, new Text({}, 'This document is read-only protected.')),
			new Paragraph({}, new Text({}, 'You cannot edit this in Word without removing protection.'))
		)
	);

	return docx;
}
`,
	},
	{
		id: 'track-changes-insertion',
		label: 'Track Changes: Insertion',
		source: `import Docx, { Insertion, Paragraph, Section, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	const date = new Date();
	const author = 'Editor';

	docx.document.set(
		new Section({},
			new Paragraph(
				{ pilcrow: { insertion: { id: 1, author, date } } },
				new Insertion({ id: 2, author, date },
					new Text({}, 'This entire paragraph was inserted.')
				)
			),
			new Paragraph({},
				new Text({}, 'This text was already here. '),
				new Insertion({ id: 3, author, date },
					new Text({}, 'This part was added later.')
				)
			)
		)
	);

	return docx;
}
`,
	},
	{
		id: 'track-changes-deletion',
		label: 'Track Changes: Deletion',
		source: `import Docx, { DeletedText, Deletion, Paragraph, Section, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	const date = new Date();
	const author = 'Editor';

	docx.document.set(
		new Section({},
			new Paragraph(
				{ pilcrow: { deletion: { id: 1, author, date } } },
				new Deletion({ id: 2, author, date },
					new DeletedText({}, 'This paragraph was deleted.')
				)
			),
			new Paragraph({},
				new Text({}, 'Existing text. '),
				new Deletion({ id: 3, author, date },
					new DeletedText({}, 'This part was removed.')
				)
			)
		)
	);

	return docx;
}
`,
	},
	{
		id: 'track-changes-formatting',
		label: 'Track Changes: Formatting',
		source: `import Docx, { Paragraph, Section, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	docx.document.set(
		new Section({},
			new Paragraph({},
				new Text(
					{
						isBold: true,
						change: { author: 'Editor', id: 1, date: new Date() },
					},
					'This text was changed to bold.'
				)
			),
			new Paragraph({},
				new Text(
					{
						isItalic: true,
						change: { author: 'Editor', id: 2, date: new Date(), isBold: true },
					},
					'This was bold, changed to italic only.'
				)
			)
		)
	);

	return docx;
}
`,
	},
	{
		id: 'track-changes-move',
		label: 'Track Changes: Move',
		source: `import Docx, { MoveFrom, MoveFromRangeStart, MoveFromRangeEnd, MoveTo, MoveToRangeStart, MoveToRangeEnd, Paragraph, Section, Text } from 'docxml';

export default function build() {
	const docx = Docx.fromNothing();

	const date = new Date();
	const author = 'Editor';

	docx.document.set(
		new Section({},
			new Paragraph(
				{ pilcrow: { moveFrom: { id: 0, author, date } } },
				new MoveFromRangeStart({ id: 1, name: 'move_0', author, date }),
				new MoveFrom({ id: 2, author, date },
					new Text({}, 'This paragraph was moved from here.')
				),
				new MoveFromRangeEnd({ id: 1 })
			),
			new Paragraph({},
				new Text({}, 'This paragraph stays in place.')
			),
			new Paragraph(
				{ pilcrow: { moveTo: { id: 3, author, date } } },
				new MoveToRangeStart({ id: 4, name: 'move_0', author, date }),
				new MoveTo({ id: 5, author, date },
					new Text({}, 'This paragraph was moved from here.')
				),
				new MoveToRangeEnd({ id: 4 })
			)
		)
	);

	return docx;
}
`,
	},
];
