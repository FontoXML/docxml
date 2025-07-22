// Top-level API
export { Docx as default } from './src/Docx.ts';

// Classes
export { type AnyComponent as DocxmlComponent } from './src/classes/src/Component.ts';

// Content components
export {
	Comment,
	type CommentChild,
	type CommentProps,
} from './src/components/comments/src/Comment.ts';
export {
	CommentRangeEnd,
	type CommentRangeEndChild,
	type CommentRangeEndProps,
} from './src/components/comments/src/CommentRangeEnd.ts';
export {
	CommentRangeStart,
	type CommentRangeStartChild,
	type CommentRangeStartProps,
} from './src/components/comments/src/CommentRangeStart.ts';
export {
	BookmarkRangeEnd,
	type BookmarkRangeEndChild,
	type BookmarkRangeEndProps,
} from './src/components/document/src/BookmarkRangeEnd.ts';
export {
	BookmarkRangeStart,
	type BookmarkRangeStartChild,
	type BookmarkRangeStartProps,
} from './src/components/document/src/BookmarkRangeStart.ts';
export {
	Break,
	type BreakChild,
	type BreakProps,
} from './src/components/document/src/Break.ts';
export {
	Cell,
	type CellChild,
	type CellProps,
} from './src/components/document/src/Cell.ts';
export {
	Field,
	type FieldChild,
	type FieldProps,
} from './src/components/document/src/Field.ts';
export {
	FieldRangeEnd,
	type FieldRangeEndChild,
	type FieldRangeEndProps,
} from './src/components/document/src/FieldRangeEnd.ts';
export {
	FieldRangeInstruction,
	type FieldRangeInstructionChild,
	type FieldRangeInstructionProps,
} from './src/components/document/src/FieldRangeInstruction.ts';
export {
	FieldRangeSeparator,
	type FieldRangeSeparatorChild,
	type FieldRangeSeparatorProps,
} from './src/components/document/src/FieldRangeSeparator.ts';
export {
	FieldRangeStart,
	type FieldRangeStartChild,
	type FieldRangeStartProps,
} from './src/components/document/src/FieldRangeStart.ts';
export {
	FootnoteReference,
	type FootnoteProps,
	type FootnoteReferenceProps,
} from './src/components/document/src/FootnoteReference.ts';
export {
	Hyperlink,
	type HyperlinkChild,
	type HyperlinkProps,
} from './src/components/document/src/Hyperlink.ts';
export {
	Image,
	type ImageChild,
	type ImageProps,
} from './src/components/document/src/Image.ts';
export {
	NonBreakingHyphen,
	type NonBreakingHyphenChild,
	type NonBreakingHyphenProps,
} from './src/components/document/src/NonBreakingHyphen.ts';
export {
	Paragraph,
	type ParagraphChild,
	type ParagraphProps,
} from './src/components/document/src/Paragraph.ts';
export {
	Row,
	type RowChild,
	type RowProps,
} from './src/components/document/src/Row.ts';
export {
	Section,
	type SectionChild,
	type SectionProps,
} from './src/components/document/src/Section.ts';
export {
	Symbol,
	type SymbolChild,
	type SymbolProps,
} from './src/components/document/src/Symbol.ts';
export {
	Tab,
	type TabChild,
	type TabProps,
} from './src/components/document/src/Tab.ts';
export {
	Table,
	type TableChild,
	type TableProps,
} from './src/components/document/src/Table.ts';
export {
	Text,
	type TextChild,
	type TextProps,
} from './src/components/document/src/Text.ts';
export {
	WatermarkText,
	type WatermarkTextChild,
	type WatermarkTextProps,
} from './src/components/document/src/WatermarkText.ts';
export {
	Move,
	type MoveChild,
	type MoveProps,
} from './src/components/track-changes/src/Move.ts';
export {
	RowAddition,
	type RowAdditionChild,
	type RowAdditionProps,
} from './src/components/track-changes/src/RowAddition.ts';
export {
	RowDeletion,
	type RowDeletionChild,
	type RowDeletionProps,
} from './src/components/track-changes/src/RowDeletion.ts';
export {
	TextAddition,
	type TextAdditionChild,
	type TextAdditionProps,
} from './src/components/track-changes/src/TextAddition.ts';
export {
	TextDeletion,
	type TextDeletionChild,
	type TextDeletionProps,
} from './src/components/track-changes/src/TextDeletion.ts';
export { FileMime } from './src/enums.ts';

// Shared properties
export {
	type Border,
	type LineBorderType,
} from './src/properties/src/shared-properties.ts';

// Utility functions
export { RelationshipType } from './src/enums.ts';
export { hex, int, type Id } from './src/utilities/src/id.ts';
export { jsx } from './src/utilities/src/jsx.ts';
export {
	cm,
	emu,
	hpt,
	inch,
	opt,
	pt,
	twip,
	type Length,
} from './src/utilities/src/length.ts';

// Archive component types
export { type CommentsXml } from './src/files/src/CommentsXml.ts';
export { type ContentTypesXml } from './src/files/src/ContentTypesXml.ts';
export {
	CustomPropertyType,
	type CustomPropertiesXml,
} from './src/files/src/CustomPropertiesXml.ts';
export {
	type DocumentChild,
	type DocumentXml,
} from './src/files/src/DocumentXml.ts';
export { type FootnotesXml } from './src/files/src/FootnotesXml.ts';
export {
	type FooterXml,
	type HeaderFooterChild,
	type HeaderXml,
} from './src/files/src/HeaderFooterXml.ts';
export { type NumberingXml } from './src/files/src/NumberingXml.ts';
export { type RelationshipsXml } from './src/files/src/RelationshipsXml.ts';
export { type SettingsXml } from './src/files/src/SettingsXml.ts';
export { type StylesXml } from './src/files/src/StylesXml.ts';
