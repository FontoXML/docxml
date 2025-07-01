// Import without assignment ensures Deno does not tree-shake this component. To avoid circular
// definitions, components register themselves in a side-effect of their module.
import './Text.ts';

import {
	Component,
	type ComponentAncestor,
	type ComponentDefinition,
} from '../classes/Component.ts';
import {
	type ChangeInformation,
	getChangeInformation,
} from '../utilities/changes.ts';
import { registerComponent } from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { QNS } from '../utilities/namespaces.ts';

/**
 * A type for indicating the start of a range of moved text. In OOXML, these are self-closing tags.
 */
export type MoveToRangeStartChildren = never;

/**
 * A type describing the props accepted by {@link TextAddition}.
 */
export type MoveToRangeStartProps = ChangeInformation & {
	colFirst?: number;
	colLast?: number;
	name?: string;
};

/**
 * A component that represents a change-tracked text that was inserted.
 */
export class MoveToRangeStart extends Component<
	MoveToRangeStartProps,
	MoveToRangeStartChildren
> {
	public static override readonly children: string[] = [];
	public static override readonly mixed: boolean = false;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		return create(
			`
				element ${QNS.w}moveToRangeStart {
					attribute ${QNS.w}id { $id },
					attribute ${QNS.w}author { $author },
					attribute ${QNS.w}date { $date },
					if (exists($name)) 
					then (
						attribute ${QNS.w}name { $name }
					)
					else (), 
					if (exists($colFirst))
					then (
						attribute ${QNS.w}colFirst { $colFirst }
					),
					if (exists($colLast))
					then (
						attribute ${QNS.w}colLast { $colLast }
					)
				}
			`,
			{
				...this.props,
				date: this.props.date.toISOString(),
				children: await this.childrenToNode(ancestry),
				name: this.props.name,
				colFirst: this.props.colFirst,
				colLast: this.props.colLast,
			}
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:moveToRangeStart';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(node: Node): MoveToRangeStart {
		const props = getChangeInformation(node);
		return new MoveToRangeStart(props);
	}
}

registerComponent(MoveToRangeStart as unknown as ComponentDefinition);
