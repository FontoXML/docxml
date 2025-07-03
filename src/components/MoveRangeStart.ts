// Import without assignment ensures Deno does not tree-shake this component. To avoid circular
// definitions, components register themselves in a side-effect of their module.
import './Text.ts';

import { Component, type ComponentDefinition } from '../classes/Component.ts';
import type { ChangeInformation } from '../utilities/changes.ts';
import { registerComponent } from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { NamespaceUri, QNS } from '../utilities/namespaces.ts';

export type MoveRangeStartProps = ChangeInformation & {
	name: string;
	type: 'from' | 'to';
};

export type MoveRangeStartChild = never;

/**
 * A component that represents a change-tracked text that was inserted.
 */
export class MoveRangeStart extends Component<
	MoveRangeStartProps,
	MoveRangeStartChild
> {
	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override toNode(): Node {
		return create(
			`	
					switch ($type)
					case 'to' return 
					element ${QNS.w}moveToRangeStart {
						attribute ${QNS.w}id { $id },
						attribute ${QNS.w}author { $author },
						attribute ${QNS.w}date { $date },
						attribute ${QNS.w}name { $name }
					}
					case 'from' return 
					element ${QNS.w}moveFromRangeStart {
						attribute ${QNS.w}id { $id },
						attribute ${QNS.w}author { $author },
						attribute ${QNS.w}date { $date },
						attribute ${QNS.w}name { $name }
					}
					default return ()
			`,
			{
				...this.props,
				type: this.props.type,
				id: this.props.id,
				name: this.props.name,
				date: this.props.date.toISOString(),
			}
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return (
			node.nodeName === 'w:moveFromRangeStart' ||
			node.nodeName === 'w:moveToRangeStart'
		);
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	// static override fromNode(node: Node): MoveRangeStart {}
}

registerComponent(MoveRangeStart as unknown as ComponentDefinition);
