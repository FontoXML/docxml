// Import without assignment ensures Deno does not tree-shake this component. To avoid circular
// definitions, components register themselves in a side-effect of their module.
import './Text.ts';

import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../classes/Component.ts';
import {
	type ChangeInformation,
	getChangeInformation,
} from '../utilities/changes.ts';
import {
	createChildComponentsFromNodes,
	registerComponent,
} from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToNodes } from '../utilities/xquery.ts';
import type { Text } from './Text.ts';

/**
 * A type specifying the children of {@link Moved}.
 */
export type MoveChild = Text;

/**
 * A type describing the props accepted by {@link Move}.
 */
export type MoveProps = ChangeInformation & { type: 'to' | 'from' };
/**
 * A component that represents a change-tracked text or paragrpah that was moved.
 */
export class Move extends Component<MoveProps, MoveChild> {
	public static override readonly children: string[] = ['Text', 'Paragraph'];

	public static override readonly mixed: boolean = false;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		return create(
			`
				let $attrs := [
					attribute ${QNS.w}id { $id }, 
					attribute ${QNS.w}date { $date }, 
					attribute ${QNS.w}author { $author }
				]
				let $moveType := 
					switch ($type)
					case 'to' return element ${QNS.w}moveTo { $attrs }
					case 'from' return element ${QNS.w}moveFrom { $attrs }
					default return () 
				return (
					copy $m := $moveType 
					modify (
						for $c in array:flatten($children)
						return (
							insert node $c into $m
						)
					)
					return $m
				)
			`,
			{
				...this.props,
				type: this.props.type,
				id: this.props.id,
				date: this.props.date.toISOString(),
				author: this.props.author,
				children: await this.childrenToNode(ancestry),
			}
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:moveFrom' || node.nodeName === 'w:moveTo';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(node: Node, context: ComponentContext): Move {
		const changeProps = getChangeInformation(node);
		const type =
			node.nodeName === `${QNS.w}moveTo` ||
			node.nodeName === `moveTo` ||
			node.nodeName === `w:moveTo`
				? 'to'
				: 'from';
		// console.log(node.nodeName);
		const children = evaluateXPathToNodes(`./*[self::${QNS.w}r]`, node);
		return new Move(
			{
				author: changeProps.author,
				date: changeProps.date,
				id: changeProps.id,
				type: type,
			},
			...createChildComponentsFromNodes<MoveChild>(
				this.children,
				children,
				context
			)
		);
	}
}

registerComponent(Move);
