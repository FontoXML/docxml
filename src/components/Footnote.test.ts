import { describe, it } from "std/testing/bdd";
import { FootnoteReference } from "./Footnote.ts";import { evaluateXPathToString } from "../utilities/xquery.ts";

describe(`Test footnote references in documents`, () => { 
    
    it('Generates an ooxml footnote reference from a node', () => { 
        const footnoteReference = new FootnoteReference({id: 1}); 
        const referenceNode = footnoteReference.toNode();  
        console.log(evaluateXPathToString(`//*`, referenceNode)); 
        
    
    })
})