/**
 * @file
 * Note this file is 99% the same as RowDeletion. Please maintain both accordingly.
 */

import { Component, ComponentAncestor, ComponentDefinition } from '../classes/Component.ts';
import type { ChangeInformation } from '../utilities/changes.ts';
import { getChangeInformation } from '../utilities/changes.ts';
import { createChildComponentsFromNodes, registerComponent } from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToBoolean, evaluateXPathToFirstNode } from '../utilities/xquery.ts';
import type { RowChild } from './Row.ts';
import { createNodeFromRow, parsePropsAndChildNodes, Row } from './Row.ts';

export type RowDeletionChild = RowChild;

export type RowDeletionProps = ChangeInformation;

export class RowDeletion extends Component<RowDeletionProps, RowDeletionChild> {
	public static readonly children: string[] = Row.children;
	public static readonly mixed: boolean = Row.mixed;

	public toNode(ancestry: ComponentAncestor[]): Node {
		const node = createNodeFromRow(this, ancestry);

		let trPr = evaluateXPathToFirstNode(`./${QNS.w}trPr`, node);
		if (!trPr) {
			trPr = create(`element ${QNS.w}trPr {}`);
			node.insertBefore(trPr, node.firstChild);
		}

		trPr.insertBefore(
			create(
				`
					element ${QNS.w}del {
						attribute ${QNS.w}id { $id },
						attribute ${QNS.w}author { $author },
						attribute ${QNS.w}date { $date }
					}
				`,
				{
					...this.props,
					date: this.props.date.toISOString(),
				},
			),
			null,
		);

		return node;
	}

	static matchesNode(node: Node): boolean {
		return evaluateXPathToBoolean(
			`
				self::${QNS.w}tr and
				./${QNS.w}trPr/${QNS.w}del and
				not(./${QNS.w}trPr/${QNS.w}ins)
			`,
			node,
		);
	}

	static fromNode(node: Node): RowDeletion {
		const { children, ...rowProps } = parsePropsAndChildNodes(node);
		const changeProps = getChangeInformation(evaluateXPathToFirstNode(`./${QNS.w}trPr`, node));

		return new RowDeletion(
			{
				...rowProps,
				...changeProps,
			},
			...createChildComponentsFromNodes<RowDeletionChild>(this.children, children),
		);
	}
}

registerComponent(RowDeletion as unknown as ComponentDefinition);