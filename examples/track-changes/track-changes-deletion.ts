/** @jsx  Docx.jsx */
import { DeletedText } from '../../lib/components/track-changes/src/DeletedText.ts';
import Docx, {
	Deletion,
	Paragraph,
	Section,
	Text,
	type TextProps,
} from '../../mod.ts';

const docx = Docx.fromNothing();

docx.document.set(
	new Section(
		{},
		new Paragraph(
			{},
			new Deletion(
				{ id: 1 },
				new DeletedText({}, 'And this is my deleted text.')
			)
		),
		new Paragraph(
			{},
			new Deletion({ id: 2 }, new DeletedText({}, 'This is my text.')),
			new Text({} as TextProps, 'Hello')
		),
		new Paragraph(
			{},
			new Deletion({ id: 2 }, new DeletedText({}, 'Deleted Text'))
		)
	)
);

docx.toFile('track-changes-deletion-ts.docx');
