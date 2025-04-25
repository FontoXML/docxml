import { Archive } from '../classes/Archive.ts';
import { Paragraph } from '../components/Paragraph.ts'; 
import { UnhandledXmlFile } from '../classes/XmlFile.ts';
import { FileMime } from '../enums.ts';
import { path } from "https://deno.land/x/dnt@0.25.2/lib/transform.deps.ts";
import { ContentTypesXml, RelationshipsXml } from "../../mod.ts";

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
	public static override async fromArchive(archive: Archive, location: string): Promise<FootnotesXml> {
		const dom = await archive.readXml(location); 
		const relsLocation = `${path.dirname(location)}/_rels/document.xml.rels`;
		console.log(relsLocation); 
		if (archive.hasFile(relsLocation)) { 
			const rels = RelationshipsXml.fromArchive(archive, FootnotesXml, location)
		}
		else { 
			console.log("NOPE") 
		}		
		return new FootnotesXml(location, await archive.readText(location));
	}
}
