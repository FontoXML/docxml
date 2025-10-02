// Import without assignment ensures Deno does not tree-shake this component. To avoid circular
// definitions, components register themselves in a side-effect of their module.
import '../../document/src/Break.ts';
import '../../document/src/FieldRangeEnd.ts';
import '../../document/src/FieldRangeInstruction.ts';
import '../../document/src/FieldRangeSeparator.ts';
import '../../document/src/FieldRangeStart.ts';
import '../../document/src/Image.ts';
import '../../document/src/NonBreakingHyphen.ts';
import '../../document/src/Symbol.ts';
import '../../document/src/Tab.ts';

import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../../../classes/src/Component.ts';
import {
	type TextProperties,
	textPropertiesFromNode,
	textPropertiesToNode,
} from '../../../properties/src/text-properties.ts';
import {
	createChildComponentsFromNodes,
	registerComponent,
} from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { QNS } from '../../../utilities/src/namespaces.ts';
import { evaluateXPathToMap } from '../../../utilities/src/xquery.ts';
import type { Break } from '../../document/src/Break.ts';
import type { FieldRangeEnd } from '../../document/src/FieldRangeEnd.ts';
import type { FieldRangeInstruction } from '../../document/src/FieldRangeInstruction.ts';
import type { FieldRangeSeparator } from '../../document/src/FieldRangeSeparator.ts';
import type { FieldRangeStart } from '../../document/src/FieldRangeStart.ts';
import type { FootnoteContinuationSeparator } from '../../document/src/FootnoteContinuationSeparator.ts';
import type { FootnoteSeparator } from '../../document/src/FootnoteSeparator.ts';
import type { Image } from '../../document/src/Image.ts';
import type { NonBreakingHyphen } from '../../document/src/NonBreakingHyphen.ts';
import type { Symbol as CharSymbol } from '../../document/src/Symbol.ts';
import type { Tab } from '../../document/src/Tab.ts';

/**
 * A type describing the components accepted as children of {@link DeletedText}.
 */
export type DeletedTextChild =
	| string
	| Break
	| FieldRangeEnd
	| FieldRangeInstruction
	| FieldRangeSeparator
	| FieldRangeStart
	| FootnoteSeparator
	| FootnoteContinuationSeparator
	| Image
	| NonBreakingHyphen
	// eslint-disable-next-line @typescript-eslint/ban-types
	| CharSymbol
	| Tab;

const __brand: unique symbol = Symbol();

/**
 * A type describing the props accepted by {@link DeletedText}.
 */
export type DeletedTextProps = TextProperties;

/**
 * A component that represents deleted text.
 * Deleted text is used within Word's track changes to indicate that text content in
 * the document has been removed. `DeletedText` must have a parent {@link Deletion}.
 */
export class DeletedText extends Component<DeletedTextProps, DeletedTextChild> {
	readonly [__brand] = 'deletedText';

	public static override readonly children: string[] = [
		'Break',
		'FieldRangeEnd',
		'FieldRangeInstruction',
		'FieldRangeSeparator',
		'FieldRangeStart',
		'FootnoteSeparator',
		'FootnoteContinuationSeparator',
		'Image',
		'NonBreakingHyphen',
		'Symbol',
		'Tab',
	];
	public static override readonly mixed: boolean = true;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	public override async toNode(ancestry: ComponentAncestor[]): Promise<Node> {
		return create(
			`
				element ${QNS.w}r {
					$rpr,
					$children
				}
			`,
			{
				rpr: await textPropertiesToNode(this.props),
				children: await Promise.all(
					this.children.map((child) => {
						if (typeof child === 'string') {
							return create(
								`element ${QNS.w}delText {
									attribute xml:space { "preserve" },
									$text
								}`,
								{
									text: child,
								}
							);
						}
						return child.toNode(ancestry);
					})
				),
			}
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:r';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(
		node: Node,
		context: ComponentContext
	): DeletedText {
		const { children, rpr } = evaluateXPathToMap<{
			rpr: Node;
			children: Node[];
		}>(
			`
				map {
					"rpr": ./${QNS.w}rPr,
					"children": array{
						./(
							${QNS.w}br,
							${QNS.w}tab,
							${QNS.w}drawing,
							${QNS.w}delText/text(),
							${QNS.w}fldChar,
							${QNS.w}instrText
						)
					}
				}
			`,
			node
		);
		return new DeletedText(
			textPropertiesFromNode(rpr) as DeletedTextProps,
			...createChildComponentsFromNodes<DeletedTextChild>(
				this.children,
				children,
				context
			)
		);
	}
}

registerComponent(DeletedText);
