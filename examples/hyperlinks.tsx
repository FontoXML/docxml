/** @jsx Docx.jsx */

import Docx, {
	BookmarkRangeEnd,
	BookmarkRangeStart,
	Hyperlink,
	Paragraph,
	Section,
	Text,
} from 'docxml';

const docx = Docx.fromNothing();

const bookmark = docx.bookmarks.create();

docx.document.set([
	<Section pageOrientation={'portrait'}>
		<Paragraph>
			<Hyperlink bookmark={bookmark}>
				<Text>This is a cross-reference to the next section</Text>
			</Hyperlink>
		</Paragraph>
	</Section>,
	<Section pageOrientation={'landscape'}>
		<BookmarkRangeStart bookmark={bookmark} />
		<Paragraph>
			<Hyperlink url="https://github.com/wvbe/docxml">
				<Text>This is a hyperlink to external target "github.com/wvbe/docxml"</Text>
			</Hyperlink>
		</Paragraph>
		<BookmarkRangeEnd bookmark={bookmark} />
	</Section>,
]);

docx.toFile('hyperlinks.docx');
