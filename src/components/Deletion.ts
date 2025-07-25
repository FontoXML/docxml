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
import type { Insertion } from './Insertion.ts';
import type { Move } from './Move.ts';
import type { Text } from './Text.ts';

/**
 * A type specifying the children of {@link Deletion}.
 */
export type DeletionChild =
	| BookmarkRangeStart
	| BookmarkRangeEnd
	| CommentRangeStart
	| CommentRangeEnd
	| Text
	| Move
	| Deletion
	| Insertion;
// ToDo add MoveRange, Addition

/**
 * A type describing the props accepted by {@link Deletion}.
 */
export type DeletionProps = ChangeInformation;

/**
 *
 * Additional documentation is here:
 * 	- https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_del_topic_ID0ESZZV.html
 * 	- https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_del_topic_ID0EMM3V.html
 *  - https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_del_topic_ID0EH23V.html
 *  - https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_del_topic_ID0EOK4V.html
 */
export class Deletion extends Component<DeletionProps, DeletionChild> {
	public static override readonly children: string[] = [
		'BookmarkRangeEnd',
		'BookmarkRangeStart',
		'CommentRangeStart',
		'CommentRangeEnd',
		'Text',
		'Move',
		'Insertion',
		this.name,
	];
	// ToDo add MoveRange, Deletion

	public static override readonly mixed: boolean = false;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		return create(
			`
				element ${QNS.w}del {
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
		return node.nodeName === 'w:del';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(node: Node, context: ComponentContext): Deletion {
		const props = getChangeInformation(node);
		return new Deletion(
			props,
			...createChildComponentsFromNodes<DeletionChild>(
				this.children,
				evaluateXPathToNodes(`./${QNS.w}r`, node),
				context
			)
		);
	}
}

registerComponent(Deletion as unknown as ComponentDefinition);
