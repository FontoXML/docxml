import * as path from 'std/path';
import { Archive } from '../classes/Archive.ts';
import { NumberMap } from '../classes/NumberMap.ts';
import { XmlFile } from '../classes/XmlFile.ts';
import { Paragraph } from '../components/Paragraph.ts';
import { FileMime } from '../enums.ts';
import { create } from '../utilities/dom.ts';
import { ALL_NAMESPACE_DECLARATIONS, QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToArray } from '../utilities/xquery.ts';
import { ContentTypesXml } from './ContentTypesXml.ts';
import { RelationshipsXml } from './RelationshipsXml.ts';

type FootnoteSeparatorType = 'separator' | 'continuationSeparator' | 'normal';

export type Footnote = {
	id: number;
	content: Paragraph[];
	type: FootnoteSeparatorType;
};

export class FootnotesXml extends XmlFile {
	public static override contentType = FileMime.footnotes;
	#footnotes = new NumberMap<Footnote>();

	public override isEmpty(): boolean {
		return !this.#footnotes.size;
	}

	public add(
		content: Paragraph[] | Paragraph,
		type: FootnoteSeparatorType
	): Footnote {
		const id = this.#footnotes.getNextAvailableKey();
		const newFootnote = {
			id: id,
			content: Array.isArray(content) ? content : [content],
			type: type,
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
						attribute w:type { $footnote('type') },
						attribute w:id { $footnote('id') },
						element w:p {
							element w:pPr { 
								element w:pStyle { 
									attribute w:val { "FootnoteText" }
								}
							}, 
							element w:r { 
								element w:rPr { 
									element w:rStyle { 
										attribute w:val { "FootnoteReference" }
									}
								},
								element w:footnoteRef {}
							}, 
							element w:r { 
								for $run in array:flatten($footnote('content'))
								return array:flatten($run/*)
							}
						}
					}
				} 
			</w:footnotes>`,
			{
				footnotes: await Promise.all(
					this.#footnotes.array().map(async (footnote) => ({
						type: footnote.type,
						id: footnote.id,
						content: await Promise.all(
							footnote.content.map((p) => p.toNode([]))
						),
					}))
				),
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
							"type": @${QNS.w}type/string()
						}
					}`,
					relsDom
				).forEach((footnote) => {
					inst.add(
						footnote.content.map((f: Node) =>
							Paragraph.fromNode(f, { archive, relationships })
						),
						footnote.type
					);
				});
			}
			// console.log(inst.#footnotes);
			return inst;
		} else {
			return inst;
		}
	}
}
