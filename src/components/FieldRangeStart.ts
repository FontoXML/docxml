import { type ComponentAncestor, Component } from '../classes/Component.ts';
import { registerComponent } from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToBoolean, evaluateXPathToMap } from '../utilities/xquery.ts';

/**
 * A type describing the components accepted as children of {@link FieldRangeStart}.
 */
export type FieldRangeStartChild = never;

/**
 * A type describing the props accepted by {@link FieldRangeStart}.
 */
export type FieldRangeStartProps = {
	isDirty?: boolean;
	isLocked?: boolean;
};

/**
 * The start of a range associated with a complex field.
 */
export class FieldRangeStart extends Component<FieldRangeStartProps, FieldRangeStartChild> {
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
					attribute ${QNS.w}fldCharType { "begin" },
					if ($isDirty) then attribute ${QNS.w}dirty { "1" } else (),
					if ($isLocked) then attribute ${QNS.w}fldLock { "1" } else ()
				}
			`,
			{
				isDirty: !!this.props.isDirty,
				isLocked: !!this.props.isLocked,
			},
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static matchesNode(node: Node): boolean {
		return evaluateXPathToBoolean('self::w:fldChar and @w:fldCharType = "begin"', node);
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static fromNode(node: Node): FieldRangeStart {
		return new FieldRangeStart(
			evaluateXPathToMap<FieldRangeStartProps>(
				`
					map {
						"isDirty": docxml:st-on-off(@${QNS.w}dirty),
						"isLocked": docxml:st-on-off(@${QNS.w}fldLock)
					}
				`,
				node,
			),
		);
	}
}

registerComponent(FieldRangeStart);
