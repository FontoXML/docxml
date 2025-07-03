// Import without assignment ensures Deno does not tree-shake this component. To avoid circular
// definitions, components register themselves in a side-effect of their module.
import './Text.ts';

import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
	type ComponentDefinition,
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
import type { Paragraph } from './Paragraph.ts';
import type { Text } from './Text.ts';
/**
 * A type for indicating the start of a range of moved text. In OOXML, these are self-closing tags.
 */
export type MoveChild = Text | Paragraph;

/**
 * A type describing the props accepted by {@link TextAddition}.
 */
export type MoveProps = ChangeInformation & { type: 'to' | 'from' };
/**
 * A component that represents a change-tracked text that was inserted.
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
				return (
					switch ($type)
					case 'to' return (
						element ${QNS.w}moveTo { 
							$attrs,
							$children
						}
					)
					case 'from' return (
						element ${QNS.w}moveFrom { 
							$attrs,
							$children
						}
					)
					default return () 
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
		const type = node.nodeName === 'w:moveTo' ? 'to' : 'from';
		const children = evaluateXPathToNodes(
			`./*[self::w:p or self::w:r]`,
			node
		);
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

registerComponent(Move as unknown as ComponentDefinition);
