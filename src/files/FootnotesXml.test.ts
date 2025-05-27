import { expect } from 'std/expect';
import { beforeAll, describe, it } from 'std/testing/bdd';
import Docx, { RelationshipType, Row, Table } from '../../mod.ts';
import { Cell } from '../components/Cell.ts';
import { Image } from '../components/Image.ts';
import { Paragraph } from '../components/Paragraph.ts';
import { Text } from '../components/Text.ts';
import { cm } from '../utilities/length.ts';
import { archive } from '../utilities/tests.ts';
import { FootnotesXml } from './FootnotesXml.ts';

const image = Deno.readFile('../test/spacekees.jpeg');

describe('Footnotes', () => {
	const document = Docx.fromNothing();
	const footnotes = document.document.footnotes;
	const testParagraph = new Paragraph({}, new Text({}, 'hello'));

	const testCell = new Cell(
		{},
		new Paragraph({}, new Text({}, 'This is a table cell'))
	);

	const testRow = new Row({}, testCell);

	const testTable = new Table({}, testRow);

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
		footnotesXml.add(testParagraph, 'normal');
		footnotesXml.add(testTable, 'normal');
		expect(document.document.isEmpty()).toBe(false);
		expect(footnotesXml.isEmpty()).toBe(false);
	});
});
