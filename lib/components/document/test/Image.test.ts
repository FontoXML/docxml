import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { Archive } from '../../../classes/src/Archive.ts';
import { Bookmarks } from '../../../classes/src/Bookmarks.ts';
import type { ComponentContext } from '../../../classes/src/Component.ts';
import { RelationshipType } from '../../../enums.ts';
import { RelationshipsXml } from '../../../files/src/RelationshipsXml.ts';
import { create } from '../../../utilities/src/dom.ts';
import { cm, pt } from '../../../utilities/src/length.ts';
import { NamespaceUri, QNS } from '../../../utilities/src/namespaces.ts';
import {
	evaluateXPathToBoolean,
	evaluateXPathToNumber,
	evaluateXPathToString,
} from '../../../utilities/src/xquery.ts';
import { Image } from '../src/Image.ts';

function createContext(): ComponentContext {
	return {
		archive: new Archive().addTextFile('word/media/image1.png', 'x'),
		relationships: new RelationshipsXml('word/_rels/document.xml.rels', [
			{
				id: 'rId1',
				type: RelationshipType.image,
				target: 'word/media/image1.png',
				isExternal: false,
				isBinary: true,
			},
		]),
		bookmarks: new Bookmarks(),
	};
}

describe('Image borders', () => {
	it('serializes all border options', () => {
		const node = new Image({
			data: Promise.resolve(new Uint8Array()),
			width: cm(1),
			height: cm(1),
			relationshipId: 'rId1',
			border: { width: pt(1), color: 'ff0000', type: 'dash' },
		}).toNode([]);

		expect(
			evaluateXPathToNumber(`descendant::${QNS.a}ln/@w/number()`, node)
		).toBe(12700);
		expect(
			evaluateXPathToString(
				`descendant::${QNS.a}ln/${QNS.a}solidFill/${QNS.a}srgbClr/@val/string()`,
				node
			)
		).toBe('ff0000');
		expect(
			evaluateXPathToString(
				`descendant::${QNS.a}ln/${QNS.a}prstDash/@val/string()`,
				node
			)
		).toBe('dash');
	});

	it('falls back to a visible line for border options that are not set', () => {
		const node = new Image({
			data: Promise.resolve(new Uint8Array()),
			width: cm(1),
			height: cm(1),
			relationshipId: 'rId1',
			border: { color: '00ff00' },
		}).toNode([]);

		expect(
			evaluateXPathToNumber(`descendant::${QNS.a}ln/@w/number()`, node)
		).toBe(9525);
		expect(
			evaluateXPathToBoolean(`exists(descendant::${QNS.a}prstDash)`, node)
		).toBe(false);
		expect(
			evaluateXPathToString(
				`descendant::${QNS.a}ln/${QNS.a}solidFill/${QNS.a}srgbClr/@val/string()`,
				node
			)
		).toBe('00ff00');
	});

	it('draws a black line when only a dash type is given', () => {
		const node = new Image({
			data: Promise.resolve(new Uint8Array()),
			width: cm(1),
			height: cm(1),
			relationshipId: 'rId1',
			border: { type: 'sysDot' },
		}).toNode([]);

		expect(
			evaluateXPathToString(
				`descendant::${QNS.a}ln/${QNS.a}solidFill/${QNS.a}srgbClr/@val/string()`,
				node
			)
		).toBe('000000');
	});

	it('reserves room for the border line so that it is not clipped', () => {
		const node = new Image({
			data: Promise.resolve(new Uint8Array()),
			width: cm(1),
			height: cm(1),
			relationshipId: 'rId1',
			border: { width: pt(3) },
		}).toNode([]);

		expect(
			evaluateXPathToNumber(
				`descendant::${QNS.wp}effectExtent/@t/number()`,
				node
			)
		).toBe(38100);
	});

	it('does not serialize a line when no border is given', () => {
		const node = new Image({
			data: Promise.resolve(new Uint8Array()),
			width: cm(1),
			height: cm(1),
			relationshipId: 'rId1',
		}).toNode([]);

		expect(
			evaluateXPathToBoolean(`exists(descendant::${QNS.a}ln)`, node)
		).toBe(false);
	});

	it('parses a border from an existing document', () => {
		const image = Image.fromNode(
			create(`
				<w:drawing xmlns:w="${NamespaceUri.w}" xmlns:wp="${NamespaceUri.wp}" xmlns:a="${NamespaceUri.a}" xmlns:pic="${NamespaceUri.pic}" xmlns:r="${NamespaceUri.r}">
					<wp:inline>
						<wp:extent cx="360000" cy="360000" />
						<wp:docPr id="1" name="Title" descr="Description" />
						<a:graphic>
							<a:graphicData uri="${NamespaceUri.pic}">
								<pic:pic>
									<pic:blipFill>
										<a:blip r:embed="rId1" />
									</pic:blipFill>
									<pic:spPr>
										<a:ln w="19050">
											<a:solidFill>
												<a:srgbClr val="0000ff" />
											</a:solidFill>
											<a:prstDash val="sysDot" />
										</a:ln>
									</pic:spPr>
								</pic:pic>
							</a:graphicData>
						</a:graphic>
					</wp:inline>
				</w:drawing>`),
			createContext()
		);

		expect(image.props.border?.width?.pt).toBe(1.5);
		expect(image.props.border?.color).toBe('0000ff');
		expect(image.props.border?.type).toBe('sysDot');
	});

	it('parses no border when the image has none', () => {
		const image = Image.fromNode(
			create(`
				<w:drawing xmlns:w="${NamespaceUri.w}" xmlns:wp="${NamespaceUri.wp}" xmlns:a="${NamespaceUri.a}" xmlns:pic="${NamespaceUri.pic}" xmlns:r="${NamespaceUri.r}">
					<wp:inline>
						<wp:extent cx="360000" cy="360000" />
						<a:graphic>
							<a:graphicData uri="${NamespaceUri.pic}">
								<pic:pic>
									<pic:blipFill>
										<a:blip r:embed="rId1" />
									</pic:blipFill>
								</pic:pic>
							</a:graphicData>
						</a:graphic>
					</wp:inline>
				</w:drawing>`),
			createContext()
		);

		expect(image.props.border).toBe(null);
	});
});
