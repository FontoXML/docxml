import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Archive } from '../../../classes/src/Archive.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { create } from '../../../utilities/src/dom.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';
import { Section } from '../src/Section.ts';
import { StructuredDocument } from '../src/StructuredDocument.ts';

const emptyContext: ComponentContext = {
	archive: new Archive(),
	relationships: null,
};

describe('Section with a structured document tag', () => {
	const body = create(`
		<w:body xmlns:w="${NamespaceUri.w}">
			<w:sdt>
				<w:sdtPr>
					<w:alias w:val="My control" />
				</w:sdtPr>
				<w:sdtContent>
					<w:p />
				</w:sdtContent>
			</w:sdt>
			<w:sectPr />
		</w:body>
	`) as Element;

	const sectPr = body.getElementsByTagNameNS(NamespaceUri.w, 'sectPr')[0];
	const section = Section.fromNode(sectPr, emptyContext);

	it('parses the w:sdt sibling as a StructuredDocument child', () => {
		expect(section.children).toHaveLength(1);
		expect(section.children[0]).toBeInstanceOf(StructuredDocument);
	});
});
