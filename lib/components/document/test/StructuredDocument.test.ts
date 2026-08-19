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

describe('StructuredDocument', () => {
	describe('matchesNode', () => {
		it('matches a w:sdt node', () => {
			const node = create(`<w:sdt xmlns:w="${NamespaceUri.w}" />`);
			expect(StructuredDocument.matchesNode(node)).toBe(true);
		});

		it('does not match another node', () => {
			const node = create(`<w:p xmlns:w="${NamespaceUri.w}" />`);
			expect(StructuredDocument.matchesNode(node)).toBe(false);
		});
	});

	describe('fromNode', () => {
		it('parses props from sdtPr', () => {
			const component = StructuredDocument.fromNode(
				create(`
					<w:sdt xmlns:w="${NamespaceUri.w}" xmlns:w15="${NamespaceUri.w15}">
						<w:sdtPr>
							<w15:appearance w15:val="tags" />
							<w:alias w:val="My control" />
							<w:lock w:val="sdtContentLocked" />
							<w:tag w:val="My tag" />
						</w:sdtPr>
						<w:sdtContent />
					</w:sdt>
				`),
				emptyContext
			);
			expect(component.props.appearance).toBe('tags');
			expect(component.props.alias).toBe('My control');
			expect(component.props.lock).toBe('sdtContentLocked');
			expect(component.props.tag).toBe('My tag');
		});

		it('parses children from sdtContent', () => {
			const component = StructuredDocument.fromNode(
				create(`
					<w:sdt xmlns:w="${NamespaceUri.w}">
						<w:sdtPr />
						<w:sdtContent>
							<w:p />
							<w:p />
						</w:sdtContent>
					</w:sdt>
				`),
				emptyContext
			);
			expect(component.children).toHaveLength(2);
			expect(component.children[0]).toBeInstanceOf(Paragraph);
		});

		it('leaves props null when sdtPr is empty', () => {
			const component = StructuredDocument.fromNode(
				create(`
					<w:sdt xmlns:w="${NamespaceUri.w}">
						<w:sdtPr />
						<w:sdtContent />
					</w:sdt>
				`),
				emptyContext
			);
			expect(component.props.appearance).toBeNull();
			expect(component.props.alias).toBeNull();
			expect(component.props.lock).toBeNull();
			expect(component.props.tag).toBeNull();
		});
	});

	describe('toNode', () => {
		it('serializes props into sdtPr', async () => {
			const component = new StructuredDocument({
				appearance: 'hidden',
				alias: 'Title',
				lock: 'sdtLocked',
				tag: 'My tag',
			});
			const output = serialize(await component.toNode([]));
			expect(normalizeXml(output)).toBe(
				normalizeXml(`
					<sdt xmlns="${NamespaceUri.w}">
						<sdtPr>
							<appearance xmlns="${NamespaceUri.w15}" xmlns:ns1="${NamespaceUri.w15}" ns1:val="hidden"/>
							<alias xmlns:ns2="${NamespaceUri.w}" ns2:val="Title"/>
							<lock xmlns:ns3="${NamespaceUri.w}" ns3:val="sdtLocked"/>
							<tag xmlns:ns4="${NamespaceUri.w}" ns4:val="My tag"/>
						</sdtPr>
						<sdtContent/>
					</sdt>
				`)
			);
		});

		it('wraps children inside sdtContent', async () => {
			const component = new StructuredDocument(
				{},
				new Paragraph({}),
				new Paragraph({})
			);
			const node = (await component.toNode([])) as Element;
			expect(node.localName).toBe('sdt');
			const content = node.getElementsByTagNameNS(
				NamespaceUri.w,
				'sdtContent'
			)[0];
			expect(
				content.getElementsByTagNameNS(NamespaceUri.w, 'p')
			).toHaveLength(2);
		});

		it('omits sdtPr child elements when props are not set', async () => {
			const component = new StructuredDocument({});
			const output = serialize(await component.toNode([]));
			expect(normalizeXml(output)).toBe(
				normalizeXml(`
					<sdt xmlns="${NamespaceUri.w}">
						<sdtPr/>
						<sdtContent/>
					</sdt>
				`)
			);
		});
	});

	describe('round trip', () => {
		it('preserves props through fromNode and toNode', async () => {
			const component = StructuredDocument.fromNode(
				create(`
					<w:sdt xmlns:w="${NamespaceUri.w}" xmlns:w15="${NamespaceUri.w15}">
						<w:sdtPr>
							<w15:appearance w15:val="boundingBox" />
							<w:alias w:val="Roundtrip" />
							<w:lock w:val="unlocked" />
							<w:tag w:val="Roundtrip tag" />
						</w:sdtPr>
						<w:sdtContent>
							<w:p />
						</w:sdtContent>
					</w:sdt>
				`),
				emptyContext
			);
			const output = serialize(await component.toNode([]));
			expect(normalizeXml(output)).toBe(
				normalizeXml(`
					<sdt xmlns="${NamespaceUri.w}">
						<sdtPr>
							<appearance xmlns="${NamespaceUri.w15}" xmlns:ns1="${NamespaceUri.w15}" ns1:val="boundingBox"/>
							<alias xmlns:ns2="${NamespaceUri.w}" ns2:val="Roundtrip"/>
							<lock xmlns:ns3="${NamespaceUri.w}" ns3:val="unlocked"/>
							<tag xmlns:ns4="${NamespaceUri.w}" ns4:val="Roundtrip tag"/>
						</sdtPr>
						<sdtContent>
							<p>
								<pPr/>
							</p>
						</sdtContent>
					</sdt>
				`)
			);
		});
	});
});
