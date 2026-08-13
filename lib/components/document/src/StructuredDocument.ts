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
import type { Deletion } from '../../track-changes/src/Deletion.ts';
import type { Insertion } from '../../track-changes/src/Insertion.ts';
import type { MoveFrom } from '../../track-changes/src/MoveFrom.ts';
import type { MoveFromRangeEnd } from '../../track-changes/src/MoveFromRangeEnd.ts';
import type { MoveFromRangeStart } from '../../track-changes/src/MoveFromRangeStart.ts';
import type { MoveTo } from '../../track-changes/src/MoveTo.ts';
import type { MoveToRangeEnd } from '../../track-changes/src/MoveToRangeEnd.ts';
import type { MoveToRangeStart } from '../../track-changes/src/MoveToRangeStart.ts';
import type { BookmarkRangeEnd } from './BookmarkRangeEnd.ts';
import type { BookmarkRangeStart } from './BookmarkRangeStart.ts';
import type { Paragraph } from './Paragraph.ts';
import type { Table } from './Table.ts';

/**
 * A type describing the components accepted as children of {@link }.
 */
export type StructuredDocumentChild =
	| Paragraph
	| Table
	| BookmarkRangeStart
	| BookmarkRangeEnd
	| MoveTo
	| MoveFrom
	| MoveToRangeStart
	| MoveToRangeEnd
	| MoveFromRangeStart
	| MoveFromRangeEnd
	| Insertion
	| Deletion
	| StructuredDocument;

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
	/**
	 * Friendly name associated with the current structured document tag.
	 */
	alias?: string | null;
	/**
	 * Set of behaviors which shall be applied to the contents of the parent structured document tag
	 * when the contents of this documents are edited
	 *
	 * @see https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_lock_topic_ID0EN6PS.html#topic_ID0EN6PS
	 */
	lock?: // Contents Cannot Be Edited At Runtime
		| 'contentLocked'
		// Contents Cannot Be Edited At Runtime And SDT Cannot Be Deleted
		| 'sdtContentLocked'
		// SDT Cannot Be Deleted
		| 'sdtLocked'
		// No Locking.  Used by Word by default.
		| 'unlocked'
		| null;
};

/**
 *
 */
export class StructuredDocument extends Component<
	StructuredDocumentProps,
	StructuredDocumentChild
> {
	public static override readonly children: string[] = [
		'Paragraph',
		'Table',
		'BookmarkRangeEnd',
		'BookmarkRangeStart',
		'MoveTo',
		'MoveFrom',
		'MoveToRangeStart',
		'MoveToRangeEnd',
		'MoveFromRangeStart',
		'MoveFromRangeEnd',
		'Insertion',
		'Deletion',
		'StructuredDocument',
	];
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
                    } else (),
                    if (exists($alias)) then element ${QNS.w}alias {
                        attribute ${QNS.w}val { $alias }
                    } else (),
                    if (exists($lock)) then element ${QNS.w}lock {
                        attribute ${QNS.w}val { $lock }
                    } else ()
                }
            `,
			{
				appearance: this.props.appearance || null,
				alias: this.props.alias || null,
				lock: this.props.lock || null,
			}
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
