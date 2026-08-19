import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Archive } from '../../../classes/src/Archive.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { create, serialize } from '../../../utilities/src/dom.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';
import { normalizeXml } from '../../../utilities/src/tests.ts';
import { Paragraph } from '../src/Paragraph.ts';
import { StructuredDocument } from '../src/StructuredDocument.ts';

const emptyContext: ComponentContext = {
	archive: new Archive(),
	relationships: null,
};

describe('Paragraph from XML', () => {
	const paragraph = Paragraph.fromNode(
		create(`
			<w:p xmlns:w="${NamespaceUri.w}" xmlns:w14="${NamespaceUri.w14}" w14:paraId="4CE0D358" w14:textId="77777777" w:rsidR="00A26C11" w:rsidRPr="00A26C11" w:rsidRDefault="00A26C11">
				<w:pPr>
					<w:pStyle w:val="Header" />
					<w:rPr>
						<w:lang w:val="en-GB" />
					</w:rPr>
				</w:pPr>
				<w:r>
					<w:rPr>
						<w:lang w:val="nl-NL" />
					</w:rPr>
					<w:t>My custom template</w:t>
				</w:r>
			</w:p>
		`),
		emptyContext
	);

	it('parses props correctly', () => {
		expect(paragraph.props.style).toBe('Header');
		expect(paragraph.props.pilcrow?.language).toBe('en-GB');
	});

	it('parses children correctly', () => {
		expect(paragraph.children).toHaveLength(1);
	});

	it('serializes correctly', async () => {
		expect(serialize(await paragraph.toNode([]))).toBe(
			normalizeXml(`
			<p xmlns="${NamespaceUri.w}" xmlns:ns1="${NamespaceUri.w14}" ns1:paraId="4CE0D358">
				<pPr>
					<pStyle xmlns:ns2="${NamespaceUri.w}" ns2:val="Header"/>
					<rPr>
						<lang xmlns:ns3="${NamespaceUri.w}" ns3:val="en-GB"/>
					</rPr>
				</pPr>
				<r>
					<rPr>
						<lang xmlns:ns4="${NamespaceUri.w}" ns4:val="nl-NL"/>
					</rPr>
					<t xml:space="preserve">My custom template</t>
				</r>
			</p>
			`)
		);
	});
});

describe('Paragraph with style change', () => {
	const now = new Date('2022-01-01');
	const paragraph = new Paragraph({
		style: 'StyleNew',
		change: {
			author: 'Wybe',
			date: now,
			id: 0,
			style: 'StyleOld',
		},
	});
	it('serializes correctly', async () => {
		expect(serialize(await paragraph.toNode([]))).toBe(
			normalizeXml(`
				<p xmlns="${NamespaceUri.w}">
					<pPr>
						<pStyle xmlns:ns1="${NamespaceUri.w}" ns1:val="StyleNew"/>
						<pPrChange xmlns:ns2="${
							NamespaceUri.w
						}" ns2:id="0" ns2:date="${now.toISOString()}" ns2:author="Wybe">
							<pPr>
								<pStyle ns2:val="StyleOld"/>
							</pPr>
							</pPrChange>
						</pPr>
					</p>
				`)
		);
	});
});

describe('Paragraph with a structured document tag', () => {
	const sdtNode = create(`
		<w:p xmlns:w="${NamespaceUri.w}">
			<w:sdt>
				<w:sdtPr>
					<w:alias w:val="My control" />
				</w:sdtPr>
				<w:sdtContent>
					<w:p />
				</w:sdtContent>
			</w:sdt>
		</w:p>
	`);

	const sdtAsObject = new Paragraph(
		{},
		new StructuredDocument({ alias: 'My control' }, new Paragraph({}))
	);

	const sdtAsNode = Paragraph.fromNode(sdtNode, emptyContext);

	it('parses the w:sdt child as a StructuredDocument', () => {
		expect(sdtAsNode.children).toHaveLength(1);
		expect(sdtAsNode.children[0]).toBeInstanceOf(StructuredDocument);
	});

	it('serializes correctly', async () => {
		expect(serialize(await sdtAsObject.toNode([]))).toEqual(
			serialize(
				create(`
					<p xmlns="${NamespaceUri.w}">
						<sdt>
							<sdtPr>
								<alias xmlns:ns1="${NamespaceUri.w}" ns1:val="My control"/>
							</sdtPr>
							<sdtContent>
								<p/>
							</sdtContent>
						</sdt>
					</p>
				`)
			)
		);
	});
});
