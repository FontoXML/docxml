import { create } from "../utilities/dom.ts";
import { QNS } from "../utilities/namespaces.ts";
import { evaluateXPathToMap } from "../utilities/xquery.ts";
import { AnyComponent, Component, ComponentAncestor, ComponentContext, ComponentNodes } from "./Component.ts";
export type FootnoteReferenceProps = { 
    id: number, 
}


export class FootnoteReference extends Component<FootnoteReferenceProps> {

    public override toNode(): Node {
        return create(`
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
            id: this.props.id
        }); 
    }
    
    static override fromNode(node: Node): FootnoteReference { 
        return new FootnoteReference(evaluateXPathToMap<FootnoteReferenceProps>(
            `map { 
                "id": ./@${QNS.w}id/number()
            }`,
            node
        ))
    }
}