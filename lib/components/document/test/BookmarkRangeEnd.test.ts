import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { create, serialize } from '../../../utilities/src/dom.ts';
import { NamespaceUri } from '../../../utilities/src/namespaces.ts';
import { BookmarkRangeEnd } from '../src/BookmarkRangeEnd.ts';

describe('BookmarkRangeEnd', () => {
	it('parses id from XML', () => {
		const node = create(`
			<w:bookmarkEnd xmlns:w="${NamespaceUri.w}" w:id="3" />
		`);
		const component = BookmarkRangeEnd.fromNode(node);
		expect(component.props.id).toBe(3);
	});

	it('serializes back to XML correctly', () => {
		const node = create(`
			<w:bookmarkEnd xmlns:w="${NamespaceUri.w}" w:id="1" />
		`);
		const component = BookmarkRangeEnd.fromNode(node);
		const output = serialize(component.toNode([]));
		expect(output).toContain('id="1"');
	});

	it('parses the displaced prop from XML', () => {
		const node = create(`
			<w:bookmarkEnd xmlns:w="${NamespaceUri.w}" w:id="4" w:displacedByCustomXml="prev" />
		`);
		const component = BookmarkRangeEnd.fromNode(node);
		expect(component.props.displaced).toBe('prev');
	});

	it('serializes the displaced prop to XML', () => {
		const component = new BookmarkRangeEnd({
			id: 2,
			displaced: 'next',
		});
		const output = serialize(component.toNode([]));
		expect(output).toContain('displacedByCustomXml="next"');
	});

	it('omits the displaced attribute when not set', () => {
		const component = new BookmarkRangeEnd({ id: 3 });
		const output = serialize(component.toNode([]));
		expect(output).not.toContain('displacedByCustomXml');
	});

	it('matches a w:bookmarkEnd node', () => {
		const node = create(
			`<w:bookmarkEnd xmlns:w="${NamespaceUri.w}" w:id="1" />`
		);
		expect(BookmarkRangeEnd.matchesNode(node)).toBe(true);
	});
});
