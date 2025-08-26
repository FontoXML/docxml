import type { ChangeInformation } from '@fontoxml/docxml';
import {
	hasUncaughtExceptionCaptureCallback,
	prependOnceListener,
} from 'node:process';
import { Deletion } from '../../components/track-changes/src/Deletion.ts';
import {
	Insertion,
	type InsertionProps,
} from '../../components/track-changes/src/Insertion.ts';
import { create } from '../../utilities/src/dom.ts';
import type { Length } from '../../utilities/src/length.ts';
import { QNS } from '../../utilities/src/namespaces.ts';
import { evaluateXPathToMap } from '../../utilities/src/xquery.ts';
import { type TableProperties } from './table-properties.ts';

export type TableRowProperties = {
	/**
	 * Specifies that the current row should be repeated at the top each new page on which the table
	 * is displayed. This can be specified for multiple rows to generate a multi-row header. Note
	 * that if the row is not the first row, then the property will be ignored.
	 */
	isHeaderRow?: null | boolean;
	/**
	 * If `true`, it prevents the contents of the row from breaking across multiple pages by moving
	 * the start of the row to the start of a new page. If the contents cannot fit on a single page,
	 * the row will start on a new page and flow onto multiple pages.
	 */
	isUnsplittable?: null | boolean;
	/**
	 * The distance between cells.
	 */
	cellSpacing?: null | Length;
	/**
	 * A property used to indicate when a table row property has changed. This will appear as a tracked
	 * change in Word's track changes feature.
	 */
	change?: null | (ChangeInformation & Omit<TableRowProperties, 'change'>);
	/**
	 * A property used to indicate that this table row should be excepted from the specified
	 * table-level properties.
	 *
	 * Read more: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_tblPrEx_topic_ID0E1GNR.html#topic_ID0E1GNR
	 */
	exception?:
		| null
		| (Omit<TableProperties, 'change'> & {
				/**
				 * A property used to indicate that there have been changes made to the exceptions.
				 *
				 * Read more here: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_tblPrExChange_topic_ID0EK1UW.html
				 */
				change?:
					| null
					| (ChangeInformation & Omit<TableProperties, 'change'>);
		  });
	/**
	 * A property used to indicate when a row has been inserted.
	 *
	 * If present, the containing row element will appear as a track-change inserted row.
	 *
	 * Read more here: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ins_topic_ID0EA14V.html
	 */
	insertion?: null | InsertionProps;
	/**
	 * A property used to indicate when a row has been deleted.
	 *
	 * If present, the containing row element will appear as a track-change deleted row.
	 *
	 * Read more here: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_del_topic_ID0EH23V.html
	 */
	deletion?: null | InsertionProps;
};

type IntermediateTableRowProps = Omit<TableRowProperties, 'exception'> & {
	exception?:
		| null
		| (Omit<TableRowProperties, 'change' | 'exception'> & {
				change?: {
					id: number;
					author?: string;
					date?: Date;
					node: Node | undefined;
				};
				node: Node | undefined;
		  });
};

export function tableRowPropertiesFromNode(
	node?: Node | null
): TableRowProperties {
	const props = node
		? evaluateXPathToMap<IntermediateTableRowProps>(
				`map {
					"isHeaderRow": docxml:ct-on-off(./${QNS.w}tblHeader),
					"isUnsplittable": docxml:ct-on-off(./${QNS.w}cantSplit),
					"cellSpacing": docxml:length(${QNS.w}tblCellSpacing[not(@${QNS.w}type = 'nil')]/@${QNS.w}w, 'twip'),
					"change": ./${QNS.w}trPrChange/map {
						"id": @${QNS.w}id/number(),
						"author": @${QNS.w}author/string(),
						"date": @${QNS.w}date/string(),
						"node": ./${QNS.w}trPr
					},
					"exception": ./${QNS.w}tblPrEx/map { 
						"node": ./${QNS.w}*[not(self::${QNS.w}tblPrExChange)],
						"change": ./${QNS.w}tblPrExChange/map {
							"id": @${QNS.w}id/number(), 
							"author": @${QNS.w}author/string(), 
							"date": @${QNS.w}/date/string(), 
							"node": ./${QNS.w}tblPrEx
						}
					},
					"insertion": ./${QNS.w}ins/map {
						"id": @${QNS.w}id/number(), 
						"author": @${QNS.w}author/string(), 
						"date": @${QNS.w}date/string()
					},
					"deletion": ./${QNS.w}del/map {
						"id": @${QNS.w}id/number(), 
						"author": @${QNS.w}author/string(), 
						"date": @${QNS.w}date/string()
					}
				}`,
				node
		  )
		: {};
	// Convert the date string to a Date object.
	if (props.change) {
		props.change.date = props.change.date
			? new Date(props.change.date)
			: undefined;
	}

	if (props.insertion) {
		props.insertion.date = props.insertion.date
			? new Date(props.insertion.date)
			: undefined;
		props.insertion.author = props.insertion.author
			? props.insertion.author
			: undefined;
	}

	if (props.exception) {
		props.exception = {
			...tableRowPropertiesFromNode(props.exception.node),
			change: props.exception.change
				? {
						date: props.exception.change.date
							? new Date(props.exception.change.date)
							: undefined,
						id: props.exception.change.id,
						author: props.exception.change.author,
						...tableRowPropertiesFromNode(
							props.exception.change.node
						),
						node: undefined,
				  }
				: undefined,
			node: undefined,
		};
	}

	if (props.deletion) {
		props.deletion.date = props.deletion.date
			? new Date(props.deletion.date)
			: undefined;
		props.deletion.author = props.deletion.author
			? props.deletion.author
			: undefined;
	}

	return props as TableRowProperties;
}

export async function tableRowPropertiesToNode(
	trpr: TableRowProperties = {}
): Promise<Node | null> {
	if (!Object.keys(trpr).length) {
		return null;
	}
	return create(
		`element ${QNS.w}trPr {
			if ($isHeaderRow) then element ${QNS.w}tblHeader {} else (),
			if ($isUnsplittable) then element ${QNS.w}cantSplit {} else (),
			if (exists($cellSpacing)) then element ${QNS.w}tblCellSpacing {
				attribute ${QNS.w}w { round($cellSpacing('twip')) },
				attribute ${QNS.w}type { "dxa" }
			} else (),
			if (exists($change)) then element ${QNS.w}trPrChange { 
				attribute ${QNS.w}id { $change('id') }, 
				if ($change('author')) then attribute ${QNS.w}author { $change('author') } else (), 
				if ($change('date')) then attribute ${QNS.w}date { $change('date') } else (),
				$change('node') 
			} else (), 
			if (exists($exception)) then element ${QNS.w}tblPrEx { 
				if (exists($exception('change'))) then element ${QNS.w}tblPrExChange { 
					attribute ${QNS.w}id { $exception('change')('id')}, 
					attribute ${QNS.w}author { $exception('change')('author')}, 
					attribute ${QNS.w}date { $exception('change')('date')}, 
					element ${QNS.w}tblPrEx { 
						array:flatten($exception('change')('node'))
					}
				} else (),
				$exception('node')
			} else (),
			$insertion,
			$deletion
		}`,
		{
			isHeaderRow: trpr.isHeaderRow || false,
			isUnsplittable: trpr.isUnsplittable || false,
			cellSpacing: trpr.cellSpacing || null,
			change: trpr.change
				? {
						id: trpr.change.id,
						author: trpr.change.author
							? trpr.change.author
							: undefined,
						date: trpr.change.date
							? new Date(trpr.change.date).toISOString()
							: undefined,
						node: await tableRowPropertiesToNode(trpr.change),
				  }
				: null,
			exception: trpr.exception
				? {
						node: await tableRowPropertiesToNode(trpr.exception),
						change: trpr.exception.change
							? {
									id: trpr.exception.change.id,
									author: trpr.exception.change.author
										? trpr.exception.change.author
										: undefined,
									date: trpr.exception.change.date
										? trpr.exception.change.date.toISOString()
										: undefined,
									node: await tableRowPropertiesToNode(
										trpr.exception.change
									),
							  }
							: null,
				  }
				: null,
			insertion: trpr.insertion
				? await new Insertion(trpr.insertion).toNode([])
				: null,
			deletion: trpr.deletion
				? await new Deletion(trpr.deletion).toNode([])
				: null,
		}
	);
}
