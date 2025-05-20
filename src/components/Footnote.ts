import { Component } from '../classes/Component.ts';
import { registerComponent } from '../utilities/components.ts';
import { create } from '../utilities/dom.ts';
import { QNS } from '../utilities/namespaces.ts';
import { evaluateXPathToMap } from '../utilities/xquery.ts';

/**
 * A type that models the properties that are applied to all of a document's footnotes.
 */
export type FootnoteProps = {
	numberingFormat?:
		| 'bullet'
		| 'chicago'
		| 'decimal'
		| 'lowerRoman'
		| 'upperRoman'
		| 'lowerLetter'
		| 'upperLetter';
	position?: 'beneathText' | 'documentEnd' | 'sectionEnd' | 'pageBottom';
	restart?: 'section' | 'eachPage' | 'continuous';
	styleName?: string;
	referenceStyleName?: string;
};

export type FootnoteReferenceProps = {
	id: number;
};

export class FootnoteReference extends Component<
	FootnoteReferenceProps & FootnoteProps
> {
	public override toNode(): Node {
		return create(
			`
            element ${QNS.w}r {
                element ${QNS.w}rPr {
                    element ${QNS.w}rStyle { 
                        attribute ${QNS.w}val { $referenceStyleName }
                    }
                }, 
                element ${QNS.w}footnoteReference { 
                    attribute ${QNS.w}id { $id }
                }
            }
        `,
			{
				id: this.props.id,
				referenceStyleName: this.props.referenceStyleName,
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
