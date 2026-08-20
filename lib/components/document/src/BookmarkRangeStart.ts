import type { Bookmark } from '../../../classes/src/Bookmarks.ts';
import {
	Component,
	type ComponentAncestor,
	type ComponentContext,
} from '../../../classes/src/Component.ts';
import { registerComponent } from '../../../utilities/src/components.ts';
import { create } from '../../../utilities/src/dom.ts';
import { QNS } from '../../../utilities/src/namespaces.ts';
import { evaluateXPathToMap } from '../../../utilities/src/xquery.ts';

/**
 * A type describing the components accepted as children of {@link BookmarkRangeStart}.
 */
export type BookmarkRangeStartChild = never;

/**
 * A type describing the props accepted by {@link BookmarkRangeStart}.
 */
export type BookmarkRangeStartProps =
	| {
			bookmark: Bookmark;
			id?: never;
			name?: never;
			displaced?: 'next' | 'prev' | null;
	  }
	// Deprecate this way:
	| {
			bookmark?: never;
			id: number;
			name: string;
			displaced?: 'next' | 'prev' | null;
	  };

/**
 * The start of a range associated with a bookmark.
 */
export class BookmarkRangeStart extends Component<
	BookmarkRangeStartProps,
	BookmarkRangeStartChild
> {
	public static override readonly children: string[] = [];

	public static override readonly mixed: boolean = false;

	/**
	 * Creates an XML DOM node for this component instance.
	 */
	// eslint-disable-next-line @typescript-eslint/no-unused-vars
	public override toNode(_ancestry: ComponentAncestor[]): Node {
		return create(
			`element ${QNS.w}bookmarkStart {
				attribute ${QNS.w}id { $id },
				attribute ${QNS.w}name { $name },
				if (exists($displaced)) then attribute ${QNS.w}displacedByCustomXml { $displaced } else ()
			}`,
			{
				id: this.props.bookmark
					? this.props.bookmark.id
					: this.props.id,
				name: this.props.bookmark
					? this.props.bookmark.name
					: this.props.name,
				displaced: this.props.displaced || null,
			}
		);
	}

	/**
	 * Asserts whether or not a given XML node correlates with this component.
	 */
	static override matchesNode(node: Node): boolean {
		return node.nodeName === 'w:bookmarkStart';
	}

	/**
	 * Instantiate this component from the XML in an existing DOCX file.
	 */
	static override fromNode(
		node: Node,
		context: ComponentContext
	): BookmarkRangeStart {
		const props = evaluateXPathToMap<BookmarkRangeStartProps>(
			`map {
				"id": ./@${QNS.w}id/number(),
				"name": ./@${QNS.w}name/string(),
				"displaced": ./@${QNS.w}displacedByCustomXml/string()
			}`,
			node
		);

		if (!props.displaced) {
			props.displaced = undefined;
		}

		context.bookmarks?.registerIdentifier(props.id!, props.name);
		return new BookmarkRangeStart(props);
	}
}

registerComponent(BookmarkRangeStart);
