import { Component } from '../../../classes/src/Component.ts';
import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';

export type FieldDefinitionChild = never;

/**
 * In OOXML, Field Definitions are a fixed set of codes that exist as text inside a field instruction.
 *
 * This enum will be used to store their names as they are added. Here we can access them using the .FieldName
 * notation, and grab the corresponding string for use in our XQUF for generating the node.
 *
 */

export enum FieldNames {
	'ADDRESSBLOCK' = 'ADDRESSBLOCK',
	'ADVANCE' = 'ADVANCE',
	'ANNOTATION' = 'ANNOTATION',
	'ASK' = 'ASK',
	'AUTHOR' = 'AUTHOR',
	'AUTONUM' = 'AUTONUM',
	'AUTONUMLGL' = 'AUTONUMLGL',
	'AUTONUMOUT' = 'AUTONUMOUT',
	'AUTOTEXT' = 'AUTOTEXT',
	'AUTOTEXTLIST' = 'AUTOTEXTLIST',
	'BARCODE' = 'BARCODE',
	'BIBLIOGRAPHY' = 'BIBLIOGRAPHY',
	'BIDI' = 'BIDI',
	'CITATION' = 'CITATION',
	'COMMENTS' = 'COMMENTS',
	'COMPARE' = 'COMPARE',
	'CREATEDATE' = 'CREATEDATE',
	'DATABASE' = 'DATABASE',
	'DATE' = 'DATE',
	'DDE' = 'DDE',
	'DDEAUTO' = 'DDEAUTO',
	'DISPLAYBARCODE' = 'DISPLAYBARCODE',
	'DOCAUTO' = 'DOCAUTO',
	'DOCPROPERTY' = 'DOCPROPERTY',
	'DOCVARIABLE' = 'DOCVARIABLE',
	'DOCPART' = 'DOCPART',
	'DOCPARTBYBIB' = 'DOCPARTBYBIB',
	'EDITTIME' = 'EDITTIME',
	'EQ' = 'EQ',
	'FILENAME' = 'FILENAME',
	'FILESIZE' = 'FILESIZE',
	'FILLIN' = 'FILLIN',
	'FORMCHECKBOX' = 'FORMCHECKBOX',
	'FORMDROPDOWN' = 'FORMDROPDOWN',
	'FORMTEXT' = 'FORMTEXT',
	'GLOSSARY' = 'GLOSSARY',
	'GOTOBUTTON' = 'GOTOBUTTON',
	'GREETINGLINE' = 'GREETINGLINE',
	'HYPERLINK' = 'HYPERLINK',
	'IF' = 'IF',
	'INCLUDE' = 'INCLUDE',
	'INCLUDEPICTURE' = 'INCLUDEPICTURE',
	'INCLUDETEXT' = 'INCLUDETEXT',
	'INDEX' = 'INDEX',
	'INFO' = 'INFO',
	'KEYWORDS' = 'KEYWORDS',
	'LASTSAVEDBY' = 'LASTSAVEDBY',
	'LINK' = 'LINK',
	'LISTNUM' = 'LISTNUM',
	'MACROBUTTON' = 'MACROBUTTON',
	'MERGEBARCODE' = 'MERGEBARCODE',
	'MERGEREC' = 'MERGEREC',
	'MERGESEQ' = 'MERGESEQ',
	'NEXT' = 'NEXT',
	'NEXTIF' = 'NEXTIF',
	'NOTEREF' = 'NOTEREF',
	'NUMCHARS' = 'NUMCHARS',
	'NUMPAGES' = 'NUMPAGES',
	'NUMWORDS' = 'NUMWORDS',
	'PAGE' = 'PAGE',
	'PAGEREF' = 'PAGEREF',
	'PRINT' = 'PRINT',
	'PRINTDATE' = 'PRINTDATE',
	'QUOTE' = 'QUOTE',
	'RD' = 'RD',
	'REF' = 'REF',
	'REFNO' = 'REFNO',
	'SAVEDATE' = 'SAVEDATE',
	'SECTION' = 'SECTION',
	'SECTIONPAGES' = 'SECTIONPAGES',
	'SEQ' = 'SEQ',
	'SET' = 'SET',
	'SHAPE' = 'SHAPE',
	'SKIPIF' = 'SKIPIF',
	'STYLEREF' = 'STYLEREF',
	'SUBJECT' = 'SUBJECT',
	'SYMBOL' = 'SYMBOL',
	'TA' = 'TA',
	'TC' = 'TC',
	'TEMPLATE' = 'TEMPLATE',
	'TIME' = 'TIME',
	'TOA' = 'TOA',
	'TITLE' = 'TITLE',
	'TOC' = 'TOC',
	'USERADDRESS' = 'USERADDRESS',
	'USERINITIALS' = 'USERINITIALS',
	'USERNAME' = 'USERNAME',
	'XE' = 'XE',
	'XREF' = 'XREF',

	// Error is NOT a Field Code for MS Word, but rather a catch-all for when a user tries to use an invalid
	// code or set of options.
	'ERROR' = 'ERROR',
}
type FieldSwitch = {
	switchName: string;
	args: unknown | null;
};

export type ToCSwitches = FieldSwitch & {
	name: 'TOC';
	value: null;
	switches: [
		{
			// Use paragraphs formatted with all or the specified range of built-in heading styles as entries in the table of contents. The switch is followed by a range of levels (for example, \o "1-3") to include levels 1-3, or by \o "*" to include all levels.
			switchName: 'o';
			args:
				| {
						min: number;
						max: number;
				  }
				| '*';
		},
	];
};

/**
 * In its text, each Field Definition has its name (e.g. 'HYPERLINK', 'DATE' or 'TOC'), and typically has a value.
 * In the case of hyperlinks, this value is a string specifying the link location ("http://www.google.com").
 *
 * Each Field Definition also has a set of "Switches". In OOXML these take the form of: "\b" (or any other letter), and
 * they specify certain behaviors of the Field. Many field definitions use the same set of letters, but to represent different
 * switches. "\o" may mean completely differnt things for HYPERLINK and TOC.
 *
 * Instead of relying on users to know this, we'll define them with human-readable property names.
 *
 * Ultimately, because this is a finite cascade of fixed options, we'll specify this with a large type with a lot of
 * prescribed paths for each possible field name we implement.
 *
 */

export type FieldDefinitionProps =
	| {
			name: FieldNames.HYPERLINK;
			value: string;
			switches?: string[];
			fieldSwitches?: {
				newWindow?: boolean;
				locationInFile?: string;
				screenTip?: string;
				hyperlinkCoordinates: string;
				target: string;
			};
	  }
	| {
			name: FieldNames.TOC;
			switches?: string[];
			fieldSwitches?: {
				levels?: {
					minimum: number;
					maximum: number;
					switch: 'o';
				};
				includePageNumbers?: {
					enabled: boolean;
					switch: '\\n';
				};
				useBuiltInHeadingStyles?: {
					enabled: boolean;
					minimum: number;
					maximum: number;
					switch: 'o';
				};
				useCustomHeadingStyles?: {
					enabled: boolean;
					names: string[];
					switch: '\\t';
				};
			};
	  }
	| ({
			name: Exclude<
				FieldNames,
				FieldNames.HYPERLINK | FieldNames.TOC | FieldNames.ERROR
			>;
			value?: string;
	  } & SharedFieldProps)
	| {
			name: FieldNames.ERROR;
			value: string | null;
			originalName?: string | null;
			rawInstruction?: string | null;
			switches?: string[];
	  };

function quoteValue(value: string): string {
	if (value.length === 0 || /\s/.test(value)) {
		return `"${value.replaceAll('"', '\\"')}"`;
	}

	return value;
}

function fieldNameFromString(rawName?: string): FieldNames {
	if (!rawName) {
		return FieldNames.ERROR;
	}

	const upperName = rawName.toUpperCase();
	return FIELD_NAMES.has(upperName)
		? (upperName as FieldNames)
		: FieldNames.ERROR;
}

function tokenizeInstruction(instruction: string): string[] {
	return instruction.match(/"(?:\\.|[^"])*"|\S+/g) ?? [];
}

function serializeHyperlinkSwitches(
	props: Extract<FieldDefinitionProps, { name: FieldNames.HYPERLINK }>
): string[] {
	const switches: string[] = [];

	if (props.fieldSwitches?.newWindow) {
		switches.push('\\n');
	}

	if (props.fieldSwitches?.locationInFile) {
		switches.push('\\l', quoteValue(props.fieldSwitches.locationInFile));
	}

	if (props.fieldSwitches?.screenTip) {
		switches.push('\\o', quoteValue(props.fieldSwitches.screenTip));
	}

	if (props.fieldSwitches?.hyperlinkCoordinates) {
		switches.push(
			'\\m',
			quoteValue(props.fieldSwitches.hyperlinkCoordinates)
		);
	}

	if (props.fieldSwitches?.target) {
		switches.push('\\t', quoteValue(props.fieldSwitches.target));
	}

	if (props.switches?.length) {
		switches.push(...props.switches);
	}

	return switches;
}

function serializeTocSwitches(
	props: Extract<FieldDefinitionProps, { name: FieldNames.TOC }>
): string[] {
	const switches: string[] = [];

	if (props.fieldSwitches?.levels) {
		switches.push(
			'\\o',
			quoteValue(
				`${props.fieldSwitches.levels.minimum}-${props.fieldSwitches.levels.maximum}`
			)
		);
	}

	if (props.fieldSwitches?.includePageNumbers?.enabled === false) {
		switches.push('\\n');
	}

	if (props.fieldSwitches?.useBuiltInHeadingStyles?.enabled) {
		switches.push(
			'\\o',
			quoteValue(
				`${props.fieldSwitches.useBuiltInHeadingStyles.minimum}-${props.fieldSwitches.useBuiltInHeadingStyles.maximum}`
			)
		);
	}

	if (
		props.fieldSwitches?.useCustomHeadingStyles?.enabled &&
		props.fieldSwitches.useCustomHeadingStyles.names.length
	) {
		switches.push(
			'\\t',
			quoteValue(
				props.fieldSwitches.useCustomHeadingStyles.names.join(',')
			)
		);
	}

	if (props.switches?.length) {
		switches.push(...props.switches);
	}

	return switches;
}

export class FieldDefinition extends Component<
	FieldDefinitionProps,
	FieldDefinitionChild
> {
	public static override readonly children: never;

	public static override readonly mixed: boolean = false;

	public override toNode(): Node {
		if (this.props.name === FieldNames.ERROR && this.props.rawInstruction) {
			return create(
				`
				element node {
					text { $instruction }
				}/text()
				`,
				{
					instruction: this.props.rawInstruction,
				}
			);
		}

		const parts: string[] = [this.props.name];

		if (this.props.name === FieldNames.ERROR && this.props.originalName) {
			parts[0] = this.props.originalName;
		}

		if ('value' in this.props && this.props.value) {
			parts.push(this.props.value);
		}

		switch (this.props.name) {
			case FieldNames.HYPERLINK:
				parts.push(...serializeHyperlinkSwitches(this.props));
				break;
			case FieldNames.TOC:
				parts.push(...serializeTocSwitches(this.props));
				break;
			default:
				if ('switches' in this.props && this.props.switches?.length) {
					parts.push(...this.props.switches);
				}
		}

		return create(
			`
			element node {
				text { $instruction }
			}/text()
			`,
			{
				instruction: parts.join(' ').trim(),
			}
		);
	}

	static override fromNode(node: Node): FieldDefinition {
		const textContent = node.textContent?.trim() ?? '';
		if (!textContent) {
			return new FieldDefinition({ name: FieldNames.ERROR, value: null });
		}

		const [rawName, ...tokens] = tokenizeInstruction(textContent);
		const name = fieldNameFromString(rawName);

		if (name === FieldNames.HYPERLINK) {
			const [valueToken, ...switches] = tokens;
			return new FieldDefinition({
				name,
				value: valueToken ?? '',
				switches,
			});
		}

		if (name === FieldNames.TOC) {
			return new FieldDefinition({
				name,
				switches: tokens,
			});
		}

		if (name === FieldNames.ERROR) {
			return new FieldDefinition({
				name,
				rawInstruction: textContent,
				originalName: rawName ?? null,
				value: tokens.length ? tokens.join(' ') : null,
			});
		}

		const [valueToken, ...switches] = tokens;
		return new FieldDefinition({
			name,
			value: valueToken,
			switches,
		});
	}
}

registerComponent(FieldDefinition);
