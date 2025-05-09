import { Paragraph } from '../components/Paragraph.ts'; 
import { Archive } from "../classes/Archive.ts";
import { XmlFile } from '../classes/XmlFile.ts';
import { FileMime } from '../enums.ts';
import { NumberMap } from "../classes/NumberMap.ts";
import * as slimdom from 'slimdom'; 
import * as path from 'std/path'; 
import { evaluateXPathToArray, evaluateXPathToNodes } from "../utilities/xquery.ts";
import { ALL_NAMESPACE_DECLARATIONS, QNS } from "../utilities/namespaces.ts";
import { RelationshipsXml } from './RelationshipsXml.ts'; 
import { ContentTypesXml } from "./ContentTypesXml.ts";
import { create } from "../utilities/dom.ts";

export enum FootnoteType { 
	"separator", 
	"continuationSeparator"
}

export type Footnote = { 
	id: number; 
	content: Paragraph[],
	type: FootnoteType
}

export class FootnotesXml extends XmlFile {
	public static override contentType = FileMime.footnotes;
	#footnotes = new NumberMap<Footnote>();

	public override isEmpty(): boolean {
		return !this.#footnotes.size; 
	}

	public add(content: Paragraph[] | Paragraph, type: FootnoteType): Footnote { 
		const id = this.#footnotes.getNextAvailableKey();
		const newFootnote = {
			id: id,
			content: Array.isArray(content) ? content : [content], 
			type: type
		}; 
		this.#footnotes.set(id, newFootnote); 
		return newFootnote; 
	}

	/**
	 * 
	 *  Creates an OOXML representation of the FootnotesXml. 
	 */
	protected override async toNode(): Promise<Document> {
		return create(`
			<w:footnotes ${ALL_NAMESPACE_DECLARATIONS}>
				for $footnote in array:flatten($footnotes)
				return (
					element w:footnote {
						attribute w:type { $type }
					}
				)
			</w:footnotes>
		`,
		{
			footnotes: await Promise.all(
				this.#footnotes.array().map(async (footnote) => ({ 
					type: footnote.type,
					id: footnote.id,
					content: await Promise.all(
						(footnote.content).map((p) => p.toNode([]))
					),
				}))
			)
		},
		true);
	}
	/**
	 * Instantiate this class by looking at the DOCX XML for it.
	 */
	public static override async fromArchive(
		archive: Archive,
		location: string
	): Promise<FootnotesXml>{
		const contentType = new ContentTypesXml(location); 
		const relsLocation = `${path.dirname(location)}/${path.basename(location)}`;
		const inst = new this(location); 
		if (archive.hasFile(relsLocation)) { 
			const relsDom = await archive.readXml(location); 
			const relationships =  await RelationshipsXml.fromArchive(archive, contentType, relsLocation);
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
					inst.add(footnote.content.map((f: Node) => Paragraph.fromNode(f, {archive, relationships})), footnote.type)
				})
			}
			// console.log(inst.#footnotes);
			return inst;
		} else { 
			return inst;
		}
	}

}
