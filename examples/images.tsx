/** @jsx Docx.jsx */
import Docx, { cm, Image, Paragraph, pt, Section, Text } from '../mod.ts';

await Docx.fromJsx(
	<Section>
		<Paragraph>
			<Text>
				<Image
					data={Deno.readFile('assets/spacekees.jpeg')}
					width={cm(16)}
					height={cm(16)}
					title='Title'
					alt='Description'
				/>
			</Text>
		</Paragraph>
		<Paragraph>
			<Text>
				Next image will look different if your word processor supports
				SVG.
			</Text>
		</Paragraph>
		<Paragraph>
			<Text>
				<Image
					data={Deno.readFile('assets/spacekees.jpeg')}
					dataExtensions={{
						svg: Deno.readTextFile('assets/git.svg'),
					}}
					width={cm(16)}
					height={cm(16)}
					title='Title'
					alt='Description'
				/>
			</Text>
		</Paragraph>
		<Paragraph>
			<Text>This image has a configurable border.</Text>
		</Paragraph>
		<Paragraph>
			<Text>
				<Image
					data={Deno.readFile('assets/spacekees.jpeg')}
					width={cm(16)}
					height={cm(16)}
					title='Title'
					alt='Description'
					border={{
						width: pt(3),
						color: 'ff0000',
						type: 'dash',
					}}
				/>
			</Text>
		</Paragraph>
	</Section>
).toFile('images.docx');
