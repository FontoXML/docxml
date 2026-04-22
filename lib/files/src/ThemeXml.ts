import type { Archive } from '../../classes/src/Archive.ts';
import { XmlFile } from '../../classes/src/XmlFile.ts';
import { FileMime } from '../../enums.ts';
import { create } from '../../utilities/src/dom.ts';
import { QNS } from '../../utilities/src/namespaces.ts';
import {
	evaluateXPathToFirstNode,
	evaluateXPathToMap,
	evaluateXPathToString,
} from '../../utilities/src/xquery.ts';

/**
 * A single color definition within a theme's color scheme.
 *
 * In OOXML, theme colors can be defined either as a system color reference
 * or as an explicit sRGB hex value.
 */
export type ThemeColor = {
	/**
	 * Whether this color is a system reference or a literal sRGB hex value.
	 */
	type: 'sysClr' | 'srgbClr';
	/**
	 * The color value, a system color name (e.g. "windowText") or a hex string (e.g. "4F81BD").
	 */
	value: string;
	/**
	 * The last-used concrete RGB value of a system color (e.g. "000000"). Only present for sysClr.
	 */
	lastClr?: string;
};

/**
 * The color scheme of an OOXML theme.
 *
 * Defines the 12 semantic color slots that all themed elements in a document
 * can reference. Changing a color here propagates to every object that uses
 * the corresponding theme color.
 */
export type ColorScheme = {
	/**
	 * The name of this color scheme (required, e.g. "Office").
	 */
	name: string;
	/**
	 * Primary dark color.
	 */
	dark1: ThemeColor;
	/**
	 * Primary light color.
	 */
	light1: ThemeColor;
	/**
	 * Secondary dark color.
	 */
	dark2: ThemeColor;
	/**
	 * Secondary light color.
	 */
	light2: ThemeColor;
	/**
	 * Accent color 1.
	 */
	accent1: ThemeColor;
	/**
	 * Accent color 2.
	 */
	accent2: ThemeColor;
	/**
	 * Accent color 3.
	 */
	accent3: ThemeColor;
	/**
	 * Accent color 4.
	 */
	accent4: ThemeColor;
	/**
	 * Accent color 5.
	 */
	accent5: ThemeColor;
	/**
	 * Accent color 6.
	 */
	accent6: ThemeColor;
	/**
	 * Hyperlink color.
	 */
	hyperlink: ThemeColor;
	/**
	 * Followed (visited) hyperlink color.
	 */
	followedHyperlink: ThemeColor;
};

/**
 * A supplemental font entry (a:font / CT_SupplementalFont).
 *
 * Maps a specific script (writing system) to a typeface, allowing the theme
 * to specify different fonts for different languages.
 */
export type Font = {
	/**
	 * The script/writing system this font applies to (e.g. "Jpan", "Arab", "Hans").
	 */
	script?: string;
	/**
	 * The font family name (e.g. "MS Gothic", "Arial").
	 */
	typeface: string;
};

/**
 * A Latin-script font definition.
 *
 * Extends {@link Font} with an optional Panose-1 classification number that
 * applications can use for font substitution when the exact typeface is unavailable.
 */
export interface LatinFont extends Font {
	/**
	 * The font family name (e.g. "Calibri", "Times New Roman").
	 */
	typeface: string;
	/**
	 * The Panose system is used by ooxml and other word processors as a reference
	 * system to classify fonts based on their attributes. e.g. Family, Serif, Weight, etc.
	 * The system represents 10 attributes with a single (hex) digit for each,
	 * with each attribute separated by a 0.
	 * (e.g. "020F0302020204030204")
	 */
	panose?: string;
}

/**
 * The font scheme of an OOXML theme.
 *
 * Defines two font collections:
 * - majorFont: used for headings, titles and other display text.
 * - minorFont: used for body/paragraph text.
 *
 * Each collection specifies a Latin font and an optional list of supplemental
 * fonts for other scripts/writing systems.
 */
export type FontScheme = {
	/**
	 * The name of this font scheme (required, e.g. "Office").
	 */
	name: string;
	/**
	 * Font collection for headings and display text.
	 */
	majorFont: {
		/**
		 * The primary Latin-script font for headings.
		 */
		latinFont: LatinFont;
		/**
		 * Supplemental fonts for non-Latin scripts.
		 */
		otherFonts: Font[];
	};
	/**
	 * Font collection for body/paragraph text.
	 */
	minorFont: {
		/**
		 * The primary Latin-script font for body text.
		 */
		latinFont: LatinFont;
		/**
		 * Supplemental fonts for non-Latin scripts.
		 */
		otherFonts: Font[];
	};
};

/**
 * Maps each descriptive ColorScheme property name to its OOXML element local name.
 */
const COLOR_SCHEME_XML_MAP: ReadonlyArray<{
	prop: keyof Omit<ColorScheme, 'name'>;
	xmlName: string;
}> = [
	{ prop: 'dark1', xmlName: 'dk1' },
	{ prop: 'light1', xmlName: 'lt1' },
	{ prop: 'dark2', xmlName: 'dk2' },
	{ prop: 'light2', xmlName: 'lt2' },
	{ prop: 'accent1', xmlName: 'accent1' },
	{ prop: 'accent2', xmlName: 'accent2' },
	{ prop: 'accent3', xmlName: 'accent3' },
	{ prop: 'accent4', xmlName: 'accent4' },
	{ prop: 'accent5', xmlName: 'accent5' },
	{ prop: 'accent6', xmlName: 'accent6' },
	{ prop: 'hyperlink', xmlName: 'hlink' },
	{ prop: 'followedHyperlink', xmlName: 'folHlink' },
];

/**
 * Represents an OOXML theme file.
 */
export class ThemeXml extends XmlFile {
	public static override contentType = FileMime.theme;

	#name: string;
	#colorScheme: ColorScheme | null;
	#fontScheme: FontScheme;

	public constructor(location: string) {
		super(location);
		this.#name = '';
		this.#colorScheme = null;
		const fallbackLatinFont: LatinFont = {
			typeface: 'Times New Roman',
			panose: '020206030504020304',
		};
		this.#fontScheme = {
			name: '',
			majorFont: {
				latinFont: fallbackLatinFont,
				otherFonts: [],
			},
			minorFont: {
				latinFont: fallbackLatinFont,
				otherFonts: [],
			},
		};
	}

	/**
	 * The optional display name of the theme.
	 */
	public get name(): string {
		return this.#name;
	}

	public set name(name: string) {
		this.#name = name;
	}

	/**
	 * The color scheme, or `null` if the source XML did not contain one.
	 */
	public get colorScheme(): ColorScheme | null {
		return this.#colorScheme;
	}

	public set colorScheme(colorScheme: ColorScheme | null) {
		this.#colorScheme = colorScheme;
	}

	/**
	 * The font scheme. Always present, falls back to Times New Roman if not in the source XML.
	 */
	public get fontScheme(): FontScheme {
		return this.#fontScheme;
	}

	public set fontScheme(fontScheme: FontScheme) {
		this.#fontScheme = fontScheme;
	}

	/**
	 * The major (heading) font collection.
	 */
	public get majorFont(): {
		latinFont: LatinFont;
		otherFonts: Font[];
	} {
		return this.#fontScheme.majorFont;
	}

	public set majorFont(majorFont: {
		latinFont: LatinFont;
		otherFonts: Font[];
	}) {
		this.#fontScheme.majorFont = majorFont;
	}

	/**
	 * The minor (body) font collection.
	 */
	public get minorFont(): {
		latinFont: LatinFont;
		otherFonts: Font[];
	} {
		return this.#fontScheme.minorFont;
	}

	public set minorFont(minorFont: {
		latinFont: LatinFont;
		otherFonts: Font[];
	}) {
		this.#fontScheme.minorFont = minorFont;
	}

	public override toNode(): Document {
		const colors = this.colorScheme
			? COLOR_SCHEME_XML_MAP.map(({ prop, xmlName }) => ({
					slot: xmlName,
					type: this.colorScheme![prop].type,
					value: this.colorScheme![prop].value,
					lastClr: this.colorScheme![prop].lastClr || '',
				}))
			: [];

		const doc = create(
			`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{
				if ($themeName != '') then attribute name { $themeName } else (),
				element a:themeElements {
					if ($hasColorScheme) then
						element a:clrScheme {
							attribute name { $clrSchemeName },
							for $c in array:flatten($colors)
							return element {
								fn:QName('http://schemas.openxmlformats.org/drawingml/2006/main', concat('a:', $c('slot')))
							} {
								if ($c('type') = 'sysClr') then
									element a:sysClr {
										attribute val { $c('value') },
										if ($c('lastClr') != '') then attribute lastClr { $c('lastClr') } else ()
									}
								else
									element a:srgbClr { attribute val { $c('value') } }
							}
						}
					else (),
					element a:fontScheme {
						if ($fontSchemeName != '') then attribute name { $fontSchemeName } else (),
						element a:majorFont {
							element a:latin {
								attribute typeface { $majorFontLatinTypeface },
								if ($majorFontLatinPanose != '') then attribute panose { $majorFontLatinPanose } else ()
							},
							for $font in array:flatten($majorOtherFonts)
							return element a:font {
								attribute script { $font('script') },
								attribute typeface { $font('typeface') }
							}
						},
						element a:minorFont {
							element a:latin {
								attribute typeface { $minorFontLatinTypeface },
								if ($minorFontLatinPanose != '') then attribute panose { $minorFontLatinPanose } else ()
							},
							for $font in array:flatten($minorOtherFonts)
							return element a:font {
								attribute script { $font('script') },
								attribute typeface { $font('typeface') }
							}
						}
					}
				}
			}</a:theme>`,
			{
				themeName: this.name,
				hasColorScheme: this.colorScheme !== null,
				clrSchemeName: this.colorScheme?.name ?? '',
				colors,
				fontSchemeName: this.fontScheme.name,
				majorFontLatinTypeface:
					this.fontScheme.majorFont.latinFont.typeface,
				majorFontLatinPanose:
					this.fontScheme.majorFont.latinFont.panose || '',
				majorOtherFonts: this.fontScheme.majorFont.otherFonts,
				minorFontLatinTypeface:
					this.fontScheme.minorFont.latinFont.typeface,
				minorFontLatinPanose:
					this.fontScheme.minorFont.latinFont.panose || '',
				minorOtherFonts: this.fontScheme.minorFont.otherFonts,
			},
			true
		);

		return doc;
	}

	/**
	 * Parse an XML DOM Document into a ThemeXml instance.
	 */
	public static fromDom(dom: Document, location: string): Promise<ThemeXml> {
		const newTheme = new ThemeXml(location);

		// Theme name
		newTheme.name = evaluateXPathToString(
			`string(./${QNS.a}theme/@name)`,
			dom
		);

		// Color scheme
		const clrSchemeNode = evaluateXPathToFirstNode(
			`./${QNS.a}theme/${QNS.a}themeElements/${QNS.a}clrScheme`,
			dom
		);
		if (clrSchemeNode) {
			newTheme.colorScheme = evaluateXPathToMap<ColorScheme>(
				`./map {
					"name": string(@name),
					${COLOR_SCHEME_XML_MAP.map(
						({ prop, xmlName }) =>
							`"${prop}": ${QNS.a}${xmlName}/map {
								"type": if (${QNS.a}sysClr) then "sysClr" else "srgbClr",
								"value": string((${QNS.a}sysClr/@val, ${QNS.a}srgbClr/@val)[1]),
								"lastClr": string(${QNS.a}sysClr/@lastClr)
							}`
					).join(',')}
				}`,
				clrSchemeNode
			);
		}

		// Font scheme
		newTheme.fontScheme = evaluateXPathToMap<FontScheme>(
			`
		./${QNS.a}theme/${QNS.a}themeElements/${QNS.a}fontScheme/map {
			"name": string(@name),
			"majorFont": map {
				"latinFont": map {
					"typeface": ${QNS.a}majorFont/${QNS.a}latin/@typeface/string(),
					"panose": ${QNS.a}majorFont/${QNS.a}latin/@panose/string()
				},
				"otherFonts": array{${QNS.a}majorFont/${QNS.a}font/map { "script": @script/string(), "typeface": @typeface/string()}}
			},
			"minorFont": map {
				"latinFont": map {
					"typeface": ${QNS.a}minorFont/${QNS.a}latin/@typeface/string(),
					"panose": ${QNS.a}minorFont/${QNS.a}latin/@panose/string()
				},
				"otherFonts": array{${QNS.a}minorFont/${QNS.a}font/map { "script": @script/string(), "typeface": @typeface/string()}}
			}
		}`,
			dom
		);

		return Promise.resolve(newTheme);
	}

	/**
	 * Instantiate this class by looking at the DOCX XML for it.
	 */
	public static override async fromArchive(
		archive: Archive,
		location?: string
	): Promise<ThemeXml> {
		// If a location is supplied, use that, otherwise use the default location for theme files.
		location = location ?? 'word/theme/theme1.xml';
		const themeDocument = await archive.readXml(location);
		return this.fromDom(themeDocument!, location);
	}
}
