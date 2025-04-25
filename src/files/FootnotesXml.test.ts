import { beforeAll, describe, expect, it, run } from 'https://deno.land/x/tincan@1.0.1/mod.ts';

import { parse, serialize } from '../utilities/dom.ts';
import { FootnotesXml } from "./FootnotesXml.ts";
import { archive } from "../utilities/tests.ts";

describe('Footnotes', () => {

    let footnotes: FootnotesXml; 

    beforeAll(async () => { 
        const testArchive = await archive('test/simple.docx'); 
        footnotes = await FootnotesXml.fromArchive(testArchive, 'word/footnotes.xml')
    }); 

    const footnote = parse(`
        <footnotes xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main" /> 
    `); 

    it('Footnotes class is instantiated from an existing .docx file', () => {
        console.log(footnotes); 
    })

    it('Correctly serializes FootnotesXml as an empty file', () => { 
        expect(serialize(footnote)).toBe(
            `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />`
        ); 
    }); 

});

run();
