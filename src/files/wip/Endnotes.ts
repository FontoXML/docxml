import { Archive } from '../../classes/Archive.ts';
import { UnhandledXmlFile } from '../../classes/XmlFile.ts';
import { ContentType } from '../../enums.ts';

export class Endnotes extends UnhandledXmlFile {
	public static contentType = ContentType.endnotes;

	/**
	 * Instantiate this class by looking at the DOCX XML for it.
	 */
	public static async fromArchive(archive: Archive, location: string): Promise<Endnotes> {
		return new Endnotes(location, await archive.readText(location));
	}
}
