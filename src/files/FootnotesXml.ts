import * as path from 'std/path';
import type { Archive } from '../classes/Archive.ts';
import { NumberMap } from '../classes/NumberMap.ts';
import { XmlFile } from '../classes/XmlFile.ts';
import { Paragraph } from '../components/Paragraph.ts';
import { FileMime } from '../enums.ts';
import { create } from '../utilities/dom.ts';
import { ALL_NAMESPACE_DECLARATIONS, QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToArray } from '../utilities/xquery.ts';
import { ContentTypesXml } from './ContentTypesXml.ts';
import { RelationshipsXml } from './RelationshipsXml.ts';

export type FootnoteSeparatorType =
	| 'separator'
	| 'continuationSeparator'
	| 'normal';

export type Footnote = {
	id: number;
	content: Paragraph[];
	type: FootnoteSeparatorType;
	styleName?: string;
	referenceStyleName?: string;
};

export class FootnotesXml extends XmlFile {
	public static override contentType = FileMime.footnotes;
	#footnotes = new NumberMap<Footnote>(1);

	public override isEmpty(): boolean {
		return !this.#footnotes.size;
	}

	/**
	 * Adds a footnote to a document.
	 * @param content A `Paragraph` or array of `Paragraph` objects that comprise the content.
	 * @param type Describes the type of footnote, either as a separator or 'normal'.
	 * @param styleName The style used for the text of the footnote positioned below document's main text.
	 * @param referenceStyleName The style used for the reference mark in the body text.
	 * @returns Returns a new `Footnote`
	 */
	public add(
		content: Paragraph[] | Paragraph,
		type: FootnoteSeparatorType,
		styleName?: string,
		referenceStyleName?: string
	): Footnote {
		const id = this.#footnotes.getNextAvailableKey();
		const newFootnote: Footnote = {
			id: id,
			content: Array.isArray(content) ? content : [content],
			type: type,
			styleName: styleName,
			referenceStyleName: referenceStyleName,
		};
		this.#footnotes.set(id, newFootnote);
		return newFootnote;
	}

	/**
	 *
	 *  Creates an OOXML representation of the FootnotesXml.
	 */
	protected override async toNode(): Promise<Document> {
		return create(
			`
			<w:footnotes ${ALL_NAMESPACE_DECLARATIONS}>
				{ for $footnote in array:flatten($footnotes)
					return element w:footnote {
						if ($footnote('type') eq 'normal') then ()
						else attribute w:type { $footnote('type') },
						attribute w:id { $footnote('id') },
						if (array:size($footnote('content')) > 0)
						then (
							element w:p {
								element w:pPr { 
									element w:pStyle { 
										attribute w:val { $footnote('styleName') }
									}
								}, 
								element w:r {
									element w:rPr { 
										element w:rStyle { 
											attribute w:val { $footnote('referenceStyleName') }
										}
									},
									element w:footnoteRef {}
								}, 
								for $run in array:flatten($footnote('content'))
									return array:flatten($run/*)
							}
						)
						else if ( $footnote('type') = "separator" ) then (
							element w:p {
								element w:r {
									element w:separator {}
								}
							}
						) else (
							element w:p {
								element w:r {
									element w:continuationSeparator {}
								}
							}
						)
					}
				} 
			</w:footnotes>`,
			{
				// In Word, footnotes with IDs -1 and 0 are reserved for the elements that visually separate the footnotes
				// from the regular flow of content. We generate those here.
				footnotes: [
					{
						type: 'separator',
						id: -1,
						content: [],
					},
					{
						type: 'continuationSeparator',
						id: 0,
						content: [],
					},
					...(await Promise.all(
						this.#footnotes.array().map(async (footnote) => ({
							type: footnote.type,
							id: footnote.id,
							content: await Promise.all(
								footnote.content.map((p) => p.toNode([]))
							),
							styleName: footnote.styleName,
							referenceStyleName: footnote.referenceStyleName,
						}))
					)),
				],
			},
			true
		);
	}
	/**
	 * Instantiate this class by looking at the DOCX XML for it.
	 */
	public static override async fromArchive(
		archive: Archive,
		location: string
	): Promise<FootnotesXml> {
		const contentType = new ContentTypesXml(location);
		const relsLocation = `${path.dirname(location)}/${path.basename(
			location
		)}`;
		const inst = new this(location);
		if (archive.hasFile(relsLocation)) {
			const relsDom = await archive.readXml(location);
			const relationships = await RelationshipsXml.fromArchive(
				archive,
				contentType,
				relsLocation
			);
			if (relationships && relationships != null) {
				evaluateXPathToArray(
					`array {
						//${QNS.w}footnote/map { 
							"id" : @${QNS.w}id/number(),
							"content": array { ./${QNS.w}p }, 
							"type": @${QNS.w}type/string(),
							"styleName": ./${QNS.w}p/${QNS.w}pPr/${QNS.w}pStyle/@${QNS.w}val/string(),
							"referenceStyleName": ./${QNS.w}r/${QNS.w}rPr/${QNS.w}rStyle/@${QNS.w}val/string()
						}
					}`,
					relsDom
				).forEach((footnote) => {
					inst.add(
						footnote.content.map((f: Node) =>
							Paragraph.fromNode(f, { archive, relationships })
						),
						footnote.type,
						footnote.styleName,
						footnote.referenceStyleName
					);
				});
			}
		}
		return inst;
	}
}
