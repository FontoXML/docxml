import { expect } from 'std/expect';
import { describe } from 'std/testing/bdd';

import { FieldDefinition, type FieldNames } from '../../../../mod.ts';
import { create, serialize } from '../../../utilities/src/dom.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';

describe('FieldDefinition', () => {
	const newFieldDefNode = create(
		`<w:instrText xmlns:w="${NamespaceUri.w}">HYPERLINK http://www.github.com \\*</w:instrText>`
	);

	describe('FieldDefinition from node', () => {
		const newFieldDef = FieldDefinition.fromNode(newFieldDefNode);
		console.log(newFieldDefNode.textContent);
		expect(newFieldDef.props.name).toBe('HYPERLINK');
	});

	describe('FieldDefinition to node', () => {
		const newFieldDef = new FieldDefinition({
			name: 'HYPERLINK' as FieldNames,
			value: 'http://wwww.google.com?',
		}).toNode();

		expect(serialize(newFieldDef)).toBe(
			'HYPERLINK http://wwww.google.com? \\*'
		);
	});
});
