import { expect } from 'std/expect';
import { beforeAll, describe, it } from 'std/testing/bdd';
import Docx from '../../mod.ts';
import { Paragraph } from '../components/Paragraph.ts';
import { Text } from '../components/Text.ts';
import { archive } from '../utilities/tests.ts';
import { FootnotesXml } from './FootnotesXml.ts';

describe('Footnotes', () => {
	const document = Docx.fromNothing();
	const footnotes = document.document.footnotes;
	const testParagraph = new Paragraph({}, new Text({}, 'hello'));
	let footnotesXml: FootnotesXml;

	beforeAll(async () => {
		const testArchive = await archive('test/simple.docx');
		footnotesXml = await FootnotesXml.fromArchive(
			testArchive,
			'word/footnotes.xml'
		);
	});

	it('Newly created footnotes are empty', () => {
		expect(footnotes.isEmpty()).toBe(true);
	});

	it('Adds a footnote', () => {
		footnotesXml.add(testParagraph, 'normal');
		expect(document.document.isEmpty()).toBe(false);
	});

	it('FootnotesXml is no longer empty after adding a footnote', () => {
		expect(footnotesXml.isEmpty()).toBe(false);
	});
});
