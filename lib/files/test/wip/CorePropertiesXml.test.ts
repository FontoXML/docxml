import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { serialize } from '../../../utilities/src/dom.ts';
import { normalizeXml } from '../../../utilities/src/tests.ts';
import { CorePropertiesXml } from '../../src/wip/CorePropertiesXml.ts';

describe('CoreProperties', () => {
	it('serializes an empty instance correctly', async () => {
		const now = new Date().toISOString();
		const instance = new CorePropertiesXml('');

		expect(
			serialize(await instance.$$$toNode()).replace(/(.\d{3})(?=Z)/g, '')
		).toBe(
			// It's more chatty than the original XML, but it is not incorrect.
			// @TODO maybe report this to slimdom some time
			normalizeXml(`
				<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
					<dc:title xmlns:dc="http://purl.org/dc/elements/1.1/"/>
					<dc:subject xmlns:dc="http://purl.org/dc/elements/1.1/"/>
					<dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/"/>
					<cp:keywords/><dc:description xmlns:dc="http://purl.org/dc/elements/1.1/"/>
					<cp:lastModifiedBy/>
					<cp:revision>1</cp:revision>
					<dcterms:created xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
					<dcterms:modified xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
				</cp:coreProperties>
			`).replace(/(.\d{3})(?=Z)/g, '')
		);
	});
});
