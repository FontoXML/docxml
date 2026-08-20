import type { Bookmark } from '../../../classes/src/Bookmarks.ts';
import {
	Component,
	type ComponentAncestor,
} from '../../../classes/src/Component.ts';
import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { QNS } from '../../../utilities/src/namespaces.ts';
import { evaluateXPathToMap } from '../../../utilities/src/xquery.ts';

/**
 * A type describing the components accepted as children of {@link BookmarkRangeEnd}.
 */
export type BookmarkRangeEndChild = never;

/**
 * A type describing the props accepted by {@link BookmarkRangeEnd}.
 */
export type BookmarkRangeEndProps =
	| { bookmark: Bookmark; id?: never; displaced?: 'next' | 'prev' | null }
	// Deprecate this way:
	| { bookmark?: never; id: number; displaced?: 'next' | 'prev' | null };

/**
 * The end of a range associated with a bookmark.
 */
export class BookmarkRangeEnd extends Component<
	BookmarkRangeEndProps,
	BookmarkRangeEndChild
> {
	public static override readonly children: string[] = [];

	public static override readonly mixed: boolean = false;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	// eslint-disable-next-line @typescript-eslint/no-unused-vars
	public override toNode(_ancestry: ComponentAncestor[]): Node {
		return create(
			`element ${QNS.w}bookmarkEnd {
				attribute ${QNS.w}id { $id },
				if (exists($displaced)) then attribute ${QNS.w}displacedByCustomXml { $displaced } else ()
			}`,
			{
				id: this.props.bookmark
					? this.props.bookmark.id
					: this.props.id,
				displaced: this.props.displaced || null,
			}
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:bookmarkEnd';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(node: Node): BookmarkRangeEnd {
		const props = evaluateXPathToMap<BookmarkRangeEndProps>(
			`map {
				"id": ./@${QNS.w}id/number(),
				"displaced": ./@${QNS.w}displacedByCustomXml/string()
			}`,
			node
		);

		if (!props.displaced) {
			props.displaced = undefined;
		}

		return new BookmarkRangeEnd(props);
	}
}

registerComponent(BookmarkRangeEnd);
