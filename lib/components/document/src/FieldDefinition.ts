// type FieldDefinition = {
// 	name: string;
// 	argument: string | null;
// 	fieldSwitch: FieldSwitch | FieldSwitch[];
// };

import { Component } from '../../../classes/src/Component.ts';
import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';

export type FieldDefinitionChild = never;

export enum FieldNames {
	'HYPERLINK' = 'HYPERLINK',
	'DATE' = 'DATE',
	'ERROR' = 'ERROR',
}

export type FieldDefinitionProps =
	| {
			name: FieldNames.HYPERLINK;
			value: string;
			fieldSwitches?: {
				newWindow?: boolean;
				locationInFile?: string;
				screenTip?: string;
				hyperlinkCoordinates: string;
				target: string;
			};
	  }
	| {
			name: FieldNames.DATE;
			value: string;
	  }
	| {
			name: FieldNames.ERROR;
			value: string | null;
	  };

export class FieldDefinition extends Component<
	FieldDefinitionProps,
	FieldDefinitionChild
> {
	public static override readonly children: never;

	public static override readonly mixed: boolean = false;

	public override toNode(): Node {
		return create(
			`
			element node { normalize-space(concat($name, " ", $value)), "\\*" }/text()
			`,
			{
				name: this.props.name,
				value: this.props.value,
			}
		);
	}

	static override fromNode(node: Node): FieldDefinition {
		const textContent = node.textContent?.split(' ');
		if (textContent) {
			console.log(textContent[0]);
			console.log(Object.entries(FieldNames));
			return new FieldDefinition({
				name:
					textContent[0] in FieldNames
						? (textContent[0] as FieldNames)
						: FieldNames.ERROR,
				value: textContent[1] ?? null,
			});
		} else {
			return new FieldDefinition({ name: FieldNames.ERROR, value: null });
		}
	}
}

registerComponent(FieldDefinition);
