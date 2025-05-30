import * as path from 'std/path';
import type { Archive } from '../classes/Archive.ts';
import { NumberMap } from '../classes/NumberMap.ts';
import { XmlFileWithContentTypes } from '../classes/XmlFile.ts';
import { Paragraph } from '../components/Paragraph.ts';
import type { Table } from '../components/Table.ts';
import { FileMime } from '../enums.ts';
import { create } from '../utilities/dom.ts';
import { ALL_NAMESPACE_DECLARATIONS, QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToArray } from '../utilities/xquery.ts';
import type { ContentTypesXml } from './ContentTypesXml.ts';
import { type File, RelationshipsXml } from './RelationshipsXml.ts';

export type FootnoteSeparatorType =
	| 'separator'
	| 'continuationSeparator'
	| 'normal';

export type FootnoteChild = Paragraph | Table;

/**
 * A type describing a footnote.
 */
export type Footnote = {
	id: number;
	content: FootnoteChild[];
	type: FootnoteSeparatorType;
	style: string;
};

export class FootnotesXml extends XmlFileWithContentTypes {
	public static override contentType = FileMime.footnotes;

	#footnotes = new NumberMap<Footnote>(1);

	public readonly relationships: RelationshipsXml;

	public constructor(
		location: string,
		relationships: RelationshipsXml = new RelationshipsXml(
			`${path.dirname(location)}/_rels/${path.basename(location)}.rels`
		)
	) {
		super(location);
		this.relationships = relationships;
	}

	public override isEmpty(): boolean {
		return !this.#footnotes.size;
	}

	/**
	 * Adds a footnote to a document.
	 * @param content A `Paragraph` or array of `Paragraph` objects that comprise the content.
	 * @param style The style used for the reference mark in the body text.
	 * @returns The identifier of the new footnote.
	 */
	public add(content: FootnoteChild | FootnoteChild[], style: string) {
		const id = this.#footnotes.getNextAvailableKey();
		this.#footnotes.set(id, {
			id,
			content: Array.isArray(content) ? content : [content],
			type: 'normal',
			style,
		});
		return id;
	}

	/**
	 * Get all XmlFile instances related to this one, including self. This helps the system
	 * serialize itself back to DOCX fullly. Probably not useful for consumers of the library.
	 *
	 * By default only returns the instance itself but no other related instances.
	 */
	public override getRelated(): File[] {
		return [this, ...this.relationships.getRelated()];
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
				 	let $content := 
						switch ($footnote('type'))
						case 'separator' return (
							attribute w:type { 'separator' },
							element w:p { 
								element w:r { 
									element w:separator {}
								}
							}
						)
						case 'continuationSeparator' return (
							attribute w:type { 'continuationSeparator' },
							element w:p { 
								element w:r { 
									element w:continuationSeparator {}
								}
							}
						)
						default return (
							if (array:size($footnote("content")) = 0)
							then (
								element w:p {
									element w:r { 
										element w:rPr { 
											element w:rStyle { 
												attribute w:val { $footnote('style') }
											}
										}, 
										element w:footnoteRef {}
									}
								}
							)
							else (
								(: 
									Get the head (first item), and tail (rest).
									This allows us to check the very first element of the footnote.
									If the node is a paragraph, then Word places the footnoteRef in the same paragraph.
									Else, the reference is in a different paragraph.
										This also applies for images, MSWords requires images to be placed in paragraphs,
										but shows them in a different paragraph.
								:)
								let $head := array:head($footnote('content'))
								let $tail := array:tail($footnote('content'))
								return if ($head[self::w:p] and not($head/descendant::w:drawing))
								then (
									(: The head is a paragraph, replace it with a new paragraph, make sure to include previous the nodes and attributes. :)
									element w:p {
										$head/@*,
										element w:r { 
											element w:rPr { 
												element w:rStyle { 
													attribute w:val { $footnote('style') }
												}
											}, 
											element w:footnoteRef {}
										},
										$head/*
									},
									$tail
								) else (
									(: The first node is not a paragraph, create a paragraph for the footnoteRef. :)
									element w:p {
										element w:r { 
											element w:rPr { 
												element w:rStyle { 
													attribute w:val { $footnote('style') }
												}
											}, 
											element w:footnoteRef {}
										}
									},
									$head,
									$tail
								)
							)
						) 
					return (
						element w:footnote {  
							attribute w:id { $footnote('id') }, 
							$content 
						} 
					)
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
							...footnote,
							content: await Promise.all(
								footnote.content.map(
									async (n) => await n.toNode([])
								)
							),
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
		contentTypes: ContentTypesXml,
		location: string
	): Promise<FootnotesXml> {
		const relsLocation = `${path.dirname(location)}/_rels/${path.basename(
			location
		)}`;
		const inst = new this(location);
		if (archive.hasFile(relsLocation)) {
			const relsDom = await archive.readXml(location);
			const relationships = await RelationshipsXml.fromArchive(
				archive,
				contentTypes,
				relsLocation
			);
			if (relationships && relationships != null) {
				evaluateXPathToArray(
					`array {
						//${QNS.w}footnote/map { 
							"id" : @${QNS.w}id/number(),
							"content": array { ./* }, 
							"type": @${QNS.w}type/string(),
							"style": ./${QNS.w}p/${QNS.w}pPr/${QNS.w}pStyle/@${QNS.w}val/string()
						}
					}`,
					relsDom
				).forEach((footnote) => {
					inst.add(
						footnote.content.map((f: Node) =>
							Paragraph.fromNode(f, { archive, relationships })
						),
						footnote.style
					);
				});
			}
		}
		return inst;
	}

	/**
	 * @deprecated FOR TEST PURPOSES ONLY
	 */
	public $$$clearFootnotes(): void {
		this.#footnotes.clear();
	}
}
