import { Paragraph } from '../components/Paragraph.ts'; 
import { Archive } from "../classes/Archive.ts";
import { XmlFile } from '../classes/XmlFile.ts';
import { FileMime } from '../enums.ts';
import { NumberMap } from "../classes/NumberMap.ts";
import * as path from 'std/path'; 
import { evaluateXPathToArray } from "../utilities/xquery.ts";
import { QNS } from "../utilities/namespaces.ts";

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
	 * Instantiate this class by looking at the DOCX XML for it.
	 */
	public static override async fromArchive(
		archive: Archive,
		location: string,
	): Promise<FootnotesXml>{
		
		const relsLocation = `${path.dirname(location)}/${path.basename(location)}`; 
		const relationships = archive.hasFile(relsLocation) 
			? await archive.readXml(location) 
			: null;
			
		const footnotesFromArchive = new this(location); 
		
		if (relationships && relationships != null) { 
			evaluateXPathToArray(
				`array {
					//${QNS.w}footnote/map { 
						"id" : @${QNS.w}id/number(),
						"content": array { ./${QNS.w}p }, 
						"type": @${QNS.w}type/string()
					}
				}`,
			relationships).forEach((n) => (
				footnotesFromArchive.add(
					n.forEach((p) => {
						Paragraph.fromNode(n, {archive, relationships})
					})
				)	
			))
		}
		else {
			return footnotesFromArchive
		}
	}
}
