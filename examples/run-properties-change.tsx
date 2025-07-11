/** @jsx  Docx.jsx */
import Docx, { Paragraph, Section, Text } from '../mod.ts';
import { TextPropertiesChange } from '../src/components/RunPropertiesChange.ts';
import { TextProperties } from '../src/properties/text-properties.ts';

// Create a new .docx file with track changes enabled.
const docxFile = Docx.fromNothing().withSettings({
	isTrackChangesEnabled: true,
});

const date = new Date();

docxFile.document.styles.add({
	id: 'TestStyle',
	name: 'TestStyle',
	type: 'character',
	text: {
		isBold: true,
		isItalic: true,
		isUnderlined: true,
	},
});

const newStyleObj: TextProperties = { isSmallCaps: true };

// Alternatively, you can use JSX:
await Docx.fromJsx(
	<Section>
		<Paragraph>
			<Text>
				This is my text, and it will reflect
				<TextPropertiesChange
					updatedStyle="TestStyle"
					author="Gabe"
					id={1}
					date={date}
				>
					some style changes
				</TextPropertiesChange>
				made as part of my document. We can use either a named style
				<TextPropertiesChange
					updatedStyle={newStyleObj}
					author="Gabe"
					id={2}
					date={date}
				>
					or apply inline styles to change it.
				</TextPropertiesChange>
			</Text>
		</Paragraph>
	</Section>
).toFile('run-changes-jsx.docx');
