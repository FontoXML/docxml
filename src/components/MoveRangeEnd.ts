// Import without assignment ensures Deno does not tree-shake this component. To avoid circular
// definitions, components register themselves in a side-effect of their module.
import './Text.ts';

import { Component, type ComponentDefinition } from '../classes/Component.ts';
import type { ChangeInformation } from '../utilities/changes.ts';
import { registerComponent } from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { NamespaceUri, QNS } from '../utilities/namespaces.ts';

/**
 * A type for indicating the start of a range of moved text. In OOXML, these are self-closing tags.
 */
export type MoveRangeEndChild = never;

export type MoveRangeEndProps = Pick<ChangeInformation, 'id'> & {
	type: 'from' | 'to';
};

/**
 * A component that represents a change-tracked text that was inserted.
 */
export class MoveRangeEnd extends Component<
	MoveRangeEndProps,
	MoveRangeEndChild
> {
	public static override readonly children: string[] = [];
	public static override readonly mixed: boolean = false;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override toNode(): Node {
		return create(
			`
				switch ($type)
				case 'to' return 
				element ${QNS.w}moveToRangeEnd {
					attribute ${QNS.w}id { $id }
				}
				case 'from' return 
				element ${QNS.w}moveFromRangeEnd { 
					attribute ${QNS.w}id { $id }
				}
				default return ()
			`,
			{
				type: this.props.type,
				id: this.props.id,
			}
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:moveFromRangeStart';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	// static override fromNode(node: Node): MoveFromRangeStart {
	// 	const props = getChangeInformation(node);
	// 	return new MoveFromRangeStart(props);
	// }
}

registerComponent(MoveRangeEnd as unknown as ComponentDefinition);
