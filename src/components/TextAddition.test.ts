import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { create, serialize } from '../utilities/dom.ts';
import { NamespaceUri } from '../utilities/namespaces.ts';
import { Text } from './Text.ts';
import { TextAddition } from './TextAddition.ts';

describe('Text', () => {
	const timeStamp = new Date();

	it('serializes correctly', async () => {
		const newAddition = new TextAddition(
			{
				id: 1,
				author: 'X',
				date: timeStamp,
			},
			new Text({}, 'Hello')
		);

		const additionNode = await newAddition.toNode([]);

		const newNode = create(`<ins xmlns="${NamespaceUri.w}" xmlns:ns1="${
			NamespaceUri.w
		}" ns1:id="1" ns1:author="X" ns1:date="${timeStamp.toISOString()}">
                        <r>
                            <t xml:space="preserve">Hello</t>
                        </r>
            </ins>`);

		expect(serialize(additionNode)).toEqual(serialize(newNode));
	});
});
