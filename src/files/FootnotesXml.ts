import { Archive } from '../classes/Archive.ts';
import { Paragraph } from '../components/Paragraph.ts'; 
import { UnhandledXmlFile } from '../classes/XmlFile.ts';
import { FileMime } from '../enums.ts';
import { path } from "https://deno.land/x/dnt@0.25.2/lib/transform.deps.ts";
import { RelationshipsXml } from "../../mod.ts";
import { ContentTypesXml } from "../../mod.ts";
import { FileLocation } from "../enums.ts";

type ReferenceMarker = { 

}

type Footnote = { 
	id: number;
	style?: string;
	contents: Paragraph[] 
}

export class FootnotesXml extends UnhandledXmlFile {
	public static override contentType = FileMime.footnotes;

	/**
	 * Instantiate this class by looking at the DOCX XML for it.
	 */
	public static override async fromArchive(archive: Archive): Promise<FootnotesXml> {
		const dom = await archive.readXml(FileLocation.footnotes); 
		const relsLocation = `${path.dirname(location)}/_rels/document.xml.rels`;
		console.log(relsLocation); 
		if (archive.hasFile(relsLocation)) { 
			console.log("YES"); 
		}
		else { 
			console.log("NOPE") 
		}		
		return new FootnotesXml(location, await archive.readText(location));
	}
}
