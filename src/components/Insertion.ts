import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
	type ComponentDefinition,
} from '../classes/Component.ts';
import {
	getChangeInformation,
	type ChangeInformation,
} from '../utilities/changes.ts';
import {
	createChildComponentsFromNodes,
	registerComponent,
} from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToNodes } from '../utilities/xquery.ts';
import type { BookmarkRangeEnd } from './BookmarkRangeEnd.ts';
import type { BookmarkRangeStart } from './BookmarkRangeStart.ts';
import type { CommentRangeEnd } from './CommentRangeEnd.ts';
import type { CommentRangeStart } from './CommentRangeStart.ts';
import type { Move } from './Move.ts';
import type { Text } from './Text.ts';

/**
 * A type specifying the children of {@link Insertion}.
 */
export type InsertionChild =
	| BookmarkRangeStart
	| BookmarkRangeEnd
	| CommentRangeStart
	| CommentRangeEnd
	| Text
	| Move
	| Insertion;
// ToDo add MoveRange, Deletion

/**
 * A type describing the props accepted by {@link Insertion}.
 */
export type InsertionProps = ChangeInformation;

/**
 *
 * Additional documentation is here:
 * 	- https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ins_topic_ID0EOW6V.html
 * 	- https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ins_topic_ID0EVH6V.html
 *  - https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ins_topic_ID0EZY5V.html
 *  - https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ins_topic_ID0EA14V.html
 *  - https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ins_topic_ID0EHJ5V.html
 */
export class Insertion extends Component<InsertionProps, InsertionChild> {
	public static override readonly children: string[] = [
		'BookmarkRangeEnd',
		'BookmarkRangeStart',
		'CommentRangeStart',
		'CommentRangeEnd',
		'Text',
		'Move',
		this.name,
	];

	public static override readonly mixed: boolean = false;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		return create(
			`
				element ${QNS.w}ins {
					attribute ${QNS.w}id { $id },
					attribute ${QNS.w}author { $author },
					attribute ${QNS.w}date { $date },
					$children
				}
			`,
			{
				...this.props,
				date: this.props.date.toISOString(),
				children: await this.childrenToNode(ancestry),
			}
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:ins';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(node: Node, context: ComponentContext): Insertion {
		const props = getChangeInformation(node);
		return new Insertion(
			props,
			...createChildComponentsFromNodes<InsertionChild>(
				this.children,
				evaluateXPathToNodes(`./${QNS.w}r`, node),
				context
			)
		);
	}
}

registerComponent(Insertion as unknown as ComponentDefinition);
