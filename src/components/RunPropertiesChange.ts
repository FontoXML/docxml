// Import without assignment ensures Deno does not tree-shake this component. To avoid circular
// definitions, components register themselves in a side-effect of their module.
import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../classes/Component.ts';
import type { TextProperties } from '../properties/text-properties.ts';
import type { ChangeInformation } from '../utilities/changes.ts';
import { registerComponent } from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToMap } from '../utilities/xquery.ts';

/**
 * A type describing the components accepted as children of {@link TextPropertiesChange}.
 */
export type TextPropertiesChangeChild = never;

/**
 * A type describing the props accepted by {@link TextPropertiesChange}.
 */
export type TextPropertiesChangeProps = ChangeInformation & {
	updatedStyle: string | TextProperties;
};

/**
 * A component that represents a table.
 */
export class TextPropertiesChange extends Component<
	TextPropertiesChangeProps,
	TextPropertiesChangeChild
> {
	public static override readonly mixed: boolean = true;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override toNode(ancestry: ComponentAncestor[]): Node {
		console.log(ancestry);
		const node = create(
			`
                element ${QNS.w}r{ 
                    element ${QNS.w}rPr {
                        $
                    }
                }
            `,
			{
                parent
            }
		);
		return node;
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:tbl';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(
		node: Node,
		context: ComponentContext
	): TextPropertiesChange {
		const properties = evaluateXPathToMap(``, node);
		return new TextPropertiesChange({
			id: 1,
			author: 'Gabe',
			date: new Date(1, 2, 2020),
			updatedStyle: 'MyStyle',
		});
	}
}

registerComponent(TextPropertiesChange);
