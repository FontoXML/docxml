// Import without assignment ensures Deno does not tree-shake this component. To avoid circular
// definitions, components register themselves in a side-effect of their module.
// import './Paragraph.ts'; // ?????????????

import {
	Component,
	type ComponentAncestor,
} from '../../../classes/src/Component.ts';

import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { QNS } from '../../../utilities/src/namespaces.ts';

import type { Paragraph } from './Paragraph.ts';

/**
 * A type describing the components accepted as children of {@link }.
 */
export type StructuredDocumentChild = Paragraph;

/**
 * A type describing the props accepted by {@link }.
 */
export type StructuredDocumentProps = {
	/**
	 * Controls how the structured document tag (content control) is rendered in Word.
	 *
	 * This maps to the `w15:appearance` element (Word 2013+ extension). Possible values:
	 * - `boundingBox` — the default; shows the control's blue bounding box with a title tab.
	 * - `tags` — shows start/end tag markers around the content.
	 * - `hidden` — shows no decoration at all (no box, no tags).
	 *
	 * @see https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2013.word.sdtappearance
	 */
	appearance?: 'boundingBox' | 'tags' | 'hidden' | null;
};

/**
 *
 */
export class StructuredDocument extends Component<
	StructuredDocumentProps,
	StructuredDocumentChild
> {
	public static override readonly children: string[] = ['Paragraph'];
	public static override readonly mixed: boolean = false;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		const properties = create(
			`
                element ${QNS.w}sdtPr {
                    if (exists($appearance)) then element ${QNS.w15}appearance {
                        attribute ${QNS.w15}val { $appearance }
                    } else ()
                }
            `,
			{ appearance: this.props.appearance || null }
		);
		return create(
			`
                element ${QNS.w}sdt {
                    $stdPr,
                    element ${QNS.w}sdtContent {
                        $children
                    }
                }
            `,
			{ stdPr: properties, children: await this.childrenToNode(ancestry) }
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:sdt';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(): StructuredDocument {
		return new StructuredDocument({});
	}
}

registerComponent(StructuredDocument);
