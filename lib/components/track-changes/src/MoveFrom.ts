// Import without assignment ensures Deno does not tree-shake this component. To avoid circular
// definitions, components register themselves in a side-effect of their module.

import type {
	ComponentAncestor,
	ComponentContext,
} from '../../../classes/src/Component.ts';
import type { ChangeInformation } from '../../../utilities/src/changes.ts';
import {
	createChildComponentsFromNodes,
	registerComponent,
} from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { QNS } from '../../../utilities/src/namespaces.ts';
import { evaluateXPathToMap } from '../../../utilities/src/xquery.ts';
import type { CommentRangeEnd } from '../../comments/src/CommentRangeEnd.ts';
import type { CommentRangeStart } from '../../comments/src/CommentRangeStart.ts';
import type { BookmarkRangeEnd } from '../../document/src/BookmarkRangeEnd.ts';
import type { BookmarkRangeStart } from '../../document/src/BookmarkRangeStart.ts';
import type { FootnoteReference } from '../../document/src/FootnoteReference.ts';
import type { Text } from '../../document/src/Text.ts';
import type { Deletion } from './Deletion.ts';
import type { Insertion } from './Insertion.ts';
import type { MoveRangeEnd } from './MoveRangeEnd.ts';
import type { MoveRangeStart } from './MoveRangeStart.ts';
import { MoveTo } from './MoveTo.ts';

/**
 * A type specifying the children of {@link Move}.
 */
export type MoveFromChild =
	| BookmarkRangeStart
	| BookmarkRangeEnd
	| CommentRangeStart
	| CommentRangeEnd
	| Text
	| MoveFrom
	| MoveTo
	| MoveRangeStart
	| MoveRangeEnd
	| Insertion
	| Deletion
	| FootnoteReference;

/**
 * A type describing the props accepted by {@link MoveFrom}.
 */
export type MoveFromProps = ChangeInformation;

/**
 * A component that represents a change-tracked text or paragraph that was moved.
 *
 * If a `Move` is present outside the text-properties, then paragraphs appear as a insertion in Word.
 *
 * Additional documentation is here:
 * 	- https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EE3IW.html#topic_ID0EE3IW
 * 	- https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_moveTo_topic_ID0EXMJW.html
 */
export class MoveFrom extends MoveTo {
	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		return create(
			`
				let $attrs := [
					attribute ${QNS.w}id { $id }, 
					if ($date) then attribute ${QNS.w}date { $date } else (),
					if ($author) then attribute ${QNS.w}author { $author } else ()
				]
				return (
					element ${QNS.w}moveFrom { 
						$attrs
					}
				)
			`,
			{
				...this.props,
				date: this.props.date ? this.props.date.toISOString() : null,
				author: this.props.author ? this.props.author : null,
				children: await this.childrenToNode(ancestry),
			}
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:moveFrom';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(node: Node, context: ComponentContext): MoveFrom {
		const { children, changeProps } = evaluateXPathToMap<{
			children: Node[];
			changeProps: MoveFromProps;
		}>(
			`
			map { 
				"children": array{./(
					${QNS.w}r |
					${QNS.w}ins |
					${QNS.w}del |
					${QNS.w}commentRangeStart |
					${QNS.w}commentRangeEnd |
					${QNS.w}bookmarkStart |
					${QNS.w}bookmarkEnd | 
					${QNS.w}moveTo | 
					${QNS.w}moveFrom | 
					${QNS.w}moveToRangeStart | 
					${QNS.w}moveToRangeEnd | 
					${QNS.w}moveFromRangeStart | 
					${QNS.w}moveFromRangeEnd
				)}, 
				"changeProps": map { 
					"id": @${QNS.w}id/number(),
					"date": if (@${QNS.w}date) then @${QNS.w}date/string() else (),
					"author": if (@${QNS.w}author) then @${QNS.w}author/string() else ()
				}

			}`,
			node,
			null,
			{ nodeName: (node as Element).localName }
		);
		return new MoveFrom(
			{
				...changeProps,
				date: changeProps.date ? new Date(changeProps.date) : undefined,
				author: changeProps.author ? changeProps.author : undefined,
			},
			...createChildComponentsFromNodes<MoveFromChild>(
				this.children,
				children,
				context
			)
		);
	}
}

registerComponent(MoveFrom);
