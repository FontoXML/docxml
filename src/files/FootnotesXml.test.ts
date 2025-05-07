import { beforeAll, describe, it } from 'std/testing/bdd';
import { expect } from 'std/expect'; 
import { Paragraph } from "../components/Paragraph.ts";
import { FootnotesXml, FootnoteType } from "./FootnotesXml.ts";
import { archive } from "../utilities/tests.ts";
import { Text } from "../components/Text.ts";
import Docx from "../../mod.ts";

describe('Footnotes', () => {

    const document = Docx.fromNothing(); 
    const footnotes = document.document.footnotes; 
    const testParagraph = new Paragraph({}, new Text({}, "hello")); 
    let footnotesXml: FootnotesXml; 

    beforeAll(async () => { 
        const testArchive = await archive('test/simple.docx'); 
        footnotesXml = await FootnotesXml.fromArchive(testArchive, 'word/footnotes.xml');
    }); 

    it('Newly created footnotes are empty', () => { 
        expect(footnotes.isEmpty()).toBe(true);
    }); 

    it('Adds a footnote', () => { 
        footnotesXml.add(testParagraph, FootnoteType.continuationSeparator);
    }); 

    it('FootnotesXml is no longer empty after adding a footnote', () => {
        expect(footnotesXml.isEmpty()).toBe(false);
    }); 

    it('Footnotes can be created from archive', async () => {
        await FootnotesXml.fromArchive(await archive('test/simple.docx'), 'word/footnotes.xml'); 
    }); 

    // it('Creates a node from XML', async () => {
    //     const testXml = `
    //         <w:footnotes ${ALL_NAMESPACE_DECLARATIONS}>
    //             <w:footnote w:id="0" w:type="continuationSeparator">
    //                 <w:p>
    //                     <w:rPr>
    //                     <w:rFonts w:ascii="Open Sans" w:hAnsi="Open Sans" w:cs="Open Sans"/>
    //                     <w:color w:val="000000"/>
    //                     <w:sz w:val="21"/>
    //                     <w:szCs w:val="21"/>
    //                     <w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>
    //                     </w:rPr>
    //                     <w:t>"Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt. Neque porro quisquam est, qui dolorem ipsum quia dolor sit amet, consectetur, adipisci velit, sed quia non numquam eius modi tempora incidunt ut labore et dolore magnam aliquam quaerat voluptatem. Ut enim ad minima veniam, quis nostrum exercitationem ullam corporis suscipit laboriosam, nisi ut aliquid ex ea commodi consequatur? Quis autem vel eum iure reprehenderit qui in ea voluptate velit esse quam nihil molestiae consequatur, vel illum qui dolorem eum fugiat quo voluptas nulla pariatur?" "Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt. Neque porro quisquam est, qui dolorem ipsum quia dolor sit amet, consectetur, adipisci velit, sed quia non numquam eius modi tempora incidunt ut labore et dolore magnam aliquam quaerat voluptatem. Ut enim ad minima veniam, quis nostrum exercitationem ullam corporis suscipit laboriosam, nisi ut aliquid ex ea commodi consequatur? Quis autem vel eum iure reprehenderit qui in ea voluptate velit esse quam nihil molestiae consequatur, vel illum qui dolorem eum fugiat quo voluptas nulla pariatur?"</w:t>
    //                     </w:r>
    //                 </w:p>
    //             </w:footnote>
    //         </w:footnotes> 
    //     `; 
    //     const testParagraph: Paragraph = new Paragraph({style: 'Normal'}, new Text({}, "Hello"));
    //     const newFootnote = {
    //         id: 1, 
    //         content: [testParagraph], 
    //         type: "separator"
    //     } 
    // })



}); 
