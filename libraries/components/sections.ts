import docx from 'https://esm.sh/docx@7.3.0';

import { DocxComponent, DocxNode } from '../types.ts';
import { asDocxArray, asJsonmlArray, assertChildrenAreOnlyOfType } from '../utilities/jsx.ts';
import { ParagraphNode } from './paragraphs.ts';
import { TableNode } from './tables.ts';

export type SectionProps = Omit<docx.ISectionOptions, 'children'> & {
	// @TODO allow TableOfContents into children too
	children?: Array<ParagraphNode | TableNode>;
};

export type SectionNode = DocxNode<'Section', docx.ISectionOptions>;

/**
 * The <Section> component represents one or more pages in a MS Word document. Each section can be
 * given different properties such as size and orientation.
 *
 * More info on its options:
 *   https://docx.js.org/#/usage/sections
 */
export const Section: DocxComponent<SectionProps, SectionNode> = async ({ children, ...rest }) => {
	await assertChildrenAreOnlyOfType('Section', children, 'Paragraph', 'Table');

	return {
		type: 'Section',
		children: children || [],
		docx: {
			...rest,
			children: await asDocxArray(children),
		},
		jsonml: ['div', ...(await asJsonmlArray(children))],
	};
};
