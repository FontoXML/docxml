import { type ComponentAncestor, Component } from '../classes/Component.ts';
import { registerComponent } from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToBoolean } from '../utilities/xquery.ts';

/**
 * A type describing the components accepted as children of {@link FieldRangeEnd}.
 */
export type FieldRangeEndChild = never;

/**
 * A type describing the props accepted by {@link FieldRangeEnd}.
 */
export type FieldRangeEndProps = { [key: string]: never };

/**
 * The end of a range associated with a complex field.
 */
export class FieldRangeEnd extends Component<FieldRangeEndProps, FieldRangeEndChild> {
	public static readonly children: string[] = [];

	public static readonly mixed: boolean = false;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	// eslint-disable-next-line @typescript-eslint/no-unused-vars
	public toNode(_ancestry: ComponentAncestor[]): Node {
		return create(
			`
				element ${QNS.w}fldChar {
					attribute ${QNS.w}fldCharType { "end" }
				}
			`,
			{},
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static matchesNode(node: Node): boolean {
		return evaluateXPathToBoolean('self::w:fldChar and @w:fldCharType = "end"', node);
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static fromNode(): FieldRangeEnd {
		return new FieldRangeEnd({});
	}
}

registerComponent(FieldRangeEnd);
