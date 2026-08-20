Images can be inserted using the `Image` component. Be sure to always nest them inside `Text`, a
requirement inherited from OOXML.

The most important property is `data`, a `UInt8Array` of image data. `docxml` will use whatever
`width` and `height` you pass in, so make sure they are proportional to the original image dimensions,
or at least something you're happy with.

```tsx
/** @jsx Docx.jsx */
import Docx, { cm, Image, Paragraph, Text } from 'docxml';

await Docx.fromJsx(
	<Paragraph>
		<Text>
			<Image
				data={Deno.readFile('test/spacekees.jpeg')}
				width={cm(16)}
				height={cm(16)}
				title='Title'
				alt='Description'
			/>
		</Text>
	</Paragraph>
).toFile('images.docx');
```

## Borders

Pass a `border` prop to draw a line around the image. Every option is optional; `width` defaults to
0.75pt and `color` to black, so `border={{}}` already gives you a thin black line. The `color` is a
hexadecimal code without leading hash, and `type` is one of the DrawingML dash styles (`solid`,
`dot`, `dash`, `lgDash`, `dashDot`, `lgDashDot`, `lgDashDotDot`, `sysDash`, `sysDot`, `sysDashDot`,
`sysDashDotDot`).

```tsx
/** @jsx Docx.jsx */
import Docx, { cm, Image, Paragraph, pt, Text } from 'docxml';

await Docx.fromJsx(
	<Paragraph>
		<Text>
			<Image
				data={Deno.readFile('test/spacekees.jpeg')}
				width={cm(16)}
				height={cm(16)}
				border={{ width: pt(3), color: 'ff0000', type: 'dash' }}
			/>
		</Text>
	</Paragraph>
).toFile('images.docx');
```
