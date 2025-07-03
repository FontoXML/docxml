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
import { registerComponent } from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToString } from '../utilities/xquery.ts';
import type { Paragraph } from './Paragraph.ts';
import type { Text } from './Text.ts';
/**
 * A type for indicating the start of a range of moved text. In OOXML, these are self-closing tags.
 */
export type MoveFromChild = Text | Paragraph;

/**
 * A type describing the props accepted by {@link TextAddition}.
 */
export type MoveFromProps = ChangeInformation & {
	name: string;
};

/**
 * A component that represents a change-tracked text that was inserted.
 */
export class MoveFrom extends Component<MoveFromProps, MoveFromChild> {
	public static override readonly children: string[] = ['Text', 'Paragraph'];

	public static override readonly mixed: boolean = true;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		console.log(ancestry);
		return create(
			`
				element ${QNS.w}moveFromRangeStart {
					attribute ${QNS.w}author { $author }, 
					attribute ${QNS.w}id { $id }, 
					attribute ${QNS.w}date { $date },
					attribute ${QNS.w}name { $name } 
				}, 
				element ${QNS.w}moveFrom { 
					attribute ${QNS.w}id { $moveFromId },
					attribute ${QNS.w}date { $date }, 
					attribute ${QNS.w}author { $author },
					$children
				},
				element ${QNS.w}moveFromRangeEnd {
					attribute ${QNS.w}id { $id }
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
		return node.nodeName === 'w:moveFrom';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(node: Node, context: ComponentContext): MoveFrom {
		const changeProps = getChangeInformation(node);
		const name = evaluateXPathToString(`./@${QNS.w}name`, node);
		console.log(name);
		return new MoveFrom({
			author: changeProps.author,
			date: changeProps.date,
			id: changeProps.id,
			name: name,
		});
	}
}

registerComponent(MoveFrom as unknown as ComponentDefinition);
