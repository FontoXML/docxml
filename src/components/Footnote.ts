import { Component } from '../classes/Component.ts';
import { registerComponent } from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToMap } from '../utilities/xquery.ts';

/**
 * A type that models the properties that are applied to all of a document's footnotes.
 */
export type DocumentFootnoteProps = {
	numberingFormat?:
		| 'bullet'
		| 'chicago'
		| 'decimal'
		| 'lowerRoman'
		| 'upperRoman'
		| 'lowerLetter'
		| 'upperLetter';
	position?: 'beneathText' | 'documentEnd' | 'sectionEnd' | 'pageBottom';
	restart?: 'section' | 'page' | 'continuous';
};

export type FootnoteReferenceProps = {
	id: number;
};

export class FootnoteReference extends Component<FootnoteReferenceProps> {
	public override toNode(): Node {
		return create(
			`
            element ${QNS.w}r {
                element ${QNS.w}Pr {
                    element ${QNS.w}rStyle { 
                        attribute ${QNS.w}val { "FootnoteReference" }
                    }
                }, 
                element ${QNS.w}footnoteReference { 
                    attribute ${QNS.w}id { $id }
                }
            }
        `,
			{
				id: this.props.id,
			}
		);
	}

	static override fromNode(node: Node): FootnoteReference {
		return new FootnoteReference(
			evaluateXPathToMap<FootnoteReferenceProps>(
				`map { 
                "id": ./@${QNS.w}id/number()
            }`,
				node
			)
		);
	}
}

registerComponent(FootnoteReference);
