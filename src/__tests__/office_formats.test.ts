import { customRunJs } from './customRunJs';
import JSZip from 'jszip';
import { createReport, listCommands } from '../index';
import { BUILT_IN_COMMANDS, RunJSFunc } from '../types';
import { parseXml } from '../xml';

const namespaces = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
const paragraph = (text: string) =>
  `<a:p><a:r><a:rPr b="1"/><a:t>${text}</a:t></a:r></a:p>`;
const shape = (content: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1200" y="2400"/><a:ext cx="360000" cy="360000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${content}</p:txBody></p:sp>`;
const slide = (content: string) =>
  `<p:sld ${namespaces}><p:cSld><p:spTree>${content}</p:spTree></p:cSld></p:sld>`;

async function presentation(parts: Record<string, string>) {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>`
  );
  zip.file(
    'ppt/presentation.xml',
    `<p:presentation ${namespaces}><p:sldIdLst/></p:presentation>`
  );
  zip.file('ppt/theme/theme1.xml', '<theme>untouched {{ title }}</theme>');
  for (const [name, xml] of Object.entries(parts)) zip.file(`ppt/${name}`, xml);
  return zip.generateAsync({ type: 'uint8array' });
}

async function render(
  parts: Record<string, string>,
  data: object = {},
  runJs?: RunJSFunc
) {
  const template = await presentation(parts);
  const output = await createReport({ template, data, runJs });
  return JSZip.loadAsync(output);
}

describe.each(['default', 'custom runJs'])('PowerPoint (%s)', execution => {
  const runJs = execution === 'custom runJs' ? customRunJs : undefined;
  it('renders split commands, multiple slides, notes, layouts, and masters', async () => {
    const split =
      '<a:p><a:r><a:t>{</a:t></a:r><a:r><a:t>{ INS ti</a:t></a:r><a:r><a:t>tle }</a:t></a:r><a:r><a:t>}</a:t></a:r></a:p>';
    const parts = {
      'slides/slide1.xml': slide(shape(split)),
      'slides/slide2.xml': slide(shape(paragraph('{{ title }}'))),
      'notesSlides/notesSlide1.xml': `<p:notes ${namespaces}><p:cSld><p:spTree>${shape(
        paragraph('{{ title }}')
      )}</p:spTree></p:cSld></p:notes>`,
      'slideLayouts/slideLayout1.xml': slide(shape(paragraph('{{ title }}'))),
      'slideMasters/slideMaster1.xml': slide(shape(paragraph('{{ title }}'))),
    };
    const output = await render(parts, { title: 'Sales & growth' }, runJs);
    for (const name of Object.keys(parts)) {
      const xml = await output.file(`ppt/${name}`)!.async('string');
      expect(xml).toContain('Sales &amp; growth');
      expect(xml).not.toContain('{{');
      await expect(parseXml(xml)).resolves.toBeDefined();
    }
    expect(await output.file('ppt/theme/theme1.xml')!.async('string')).toBe(
      '<theme>untouched {{ title }}</theme>'
    );
  });

  it('executes JavaScript, aliases, paragraph loops, and conditionals', async () => {
    const text = [
      '{{ EXEC prefix = "Item: " }}',
      '{{ ALIAS item INS prefix + $item }}',
      '{{ FOR item IN items }}',
      '{{ *item }}',
      '{{ END-FOR item }}',
      '{{ IF visible }}Shown{{ END-IF }}',
      '{{ IF false }}Hidden{{ END-IF }}',
    ]
      .map(paragraph)
      .join('');
    const output = await render(
      { 'slides/slide1.xml': slide(shape(text)) },
      { items: ['A', 'B'], visible: true },
      runJs
    );
    const xml = await output.file('ppt/slides/slide1.xml')!.async('string');
    expect(xml).toContain('Item: A');
    expect(xml).toContain('Item: B');
    expect(xml).toContain('Shown');
    expect(xml).not.toContain('Hidden');
    expect(xml).not.toContain('w:');
  });

  it('places line breaks between DrawingML runs', async () => {
    const output = await render(
      { 'slides/slide1.xml': slide(shape(paragraph('{{ text }}'))) },
      { text: 'First\nSecond' },
      runJs
    );
    const xml = await output.file('ppt/slides/slide1.xml')!.async('string');
    expect(xml).toMatch(/First<\/a:t><\/a:r><a:br\/><a:r><a:t[^>]*>Second/);
    expect(xml).not.toContain('w:br');
    await expect(parseXml(xml)).resolves.toBeDefined();
  });

  it('keeps one text element per run when a tag has surrounding text', async () => {
    const output = await render(
      { 'slides/slide1.xml': slide(shape(paragraph('Hello {{ title }}!'))) },
      { title: 'World' },
      runJs
    );
    const xml = await output.file('ppt/slides/slide1.xml')!.async('string');
    expect(xml).toMatch(/<a:t[^>]*>Hello World!<\/a:t>/);
    expect(xml.match(/<a:t\b/g)).toHaveLength(1);
  });

  it('repeats table rows and removes empty loops', async () => {
    const row = (text: string) =>
      `<a:tr h="360000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraph(
        text
      )}</a:txBody><a:tcPr/></a:tc></a:tr>`;
    const table = `<p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="360000"/></a:tblGrid>${row(
      '{{ FOR item IN items }}'
    )}${row('{{ $item }}')}${row(
      '{{ END-FOR item }}'
    )}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
    for (const items of [['A', 'B'], []]) {
      const output = await render(
        { 'slides/slide1.xml': slide(table) },
        { items },
        runJs
      );
      const xml = await output.file('ppt/slides/slide1.xml')!.async('string');
      expect(xml.match(/<a:tr\b/g) ?? []).toHaveLength(items.length);
      for (const item of items) expect(xml).toContain(`>${item}</a:t>`);
      expect(xml).not.toContain('{{');
    }
  });

  it('writes pictures and hyperlinks to each slide relationship file', async () => {
    const content =
      shape(paragraph('{{ IMAGE image }}')) +
      shape(paragraph('{{ LINK link }}'));
    const output = await render(
      {
        'slides/slide1.xml': slide(content),
        'slides/slide2.xml': slide(content),
      },
      {
        image: {
          data: Buffer.from('image'),
          extension: '.png',
          width: 2,
          height: 3,
        },
        link: { url: 'https://example.com', label: 'Visit' },
      },
      runJs
    );
    for (const name of ['slide1', 'slide2']) {
      const xml = await output.file(`ppt/slides/${name}.xml`)!.async('string');
      expect(xml).toContain('<p:pic>');
      expect(xml).toContain('x="1200" y="2400"');
      expect(xml).toContain('cx="720000" cy="1080000"');
      expect(xml).toContain('<a:hlinkClick r:id="link1"/>');
      expect(xml).toContain('<a:rPr b="1">');
      expect(xml).not.toContain('w:');
      const rels = await output
        .file(`ppt/slides/_rels/${name}.xml.rels`)!
        .async('string');
      expect(rels).toContain('Target="https://example.com"');
      const target = /Target="(\.\.\/media\/[^\"]+)"/.exec(rels)![1];
      expect(output.file(`ppt/${target.slice(3)}`)).not.toBeNull();
      await expect(parseXml(xml)).resolves.toBeDefined();
    }
    expect(await output.file('[Content_Types].xml')!.async('string')).toContain(
      'ContentType="image/png"'
    );
    expect(output.file('ppt/_rels/presentation.xml.rels')).toBeNull();
  });
});

it('preserves existing relationship IDs and shape IDs', async () => {
  const template = await presentation({
    'slides/slide1.xml': slide(
      shape(paragraph('{{ IMAGE image }}')) +
        shape(paragraph('{{ LINK link }}'))
    ),
    'slides/_rels/slide1.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="img11" Type="image" Target="../media/existing.png"/><Relationship Id="link1" Type="hyperlink" Target="https://existing.example" TargetMode="External"/></Relationships>',
    'media/existing.png': 'existing image',
  });
  const output = await JSZip.loadAsync(
    await createReport({
      template,
      data: {
        image: {
          data: Buffer.from('image'),
          extension: '.png',
          width: 1,
          height: 1,
        },
        link: { url: 'https://new.example' },
      },
    })
  );
  const xml = await output.file('ppt/slides/slide1.xml')!.async('string');
  expect(xml).toContain('r:embed="img12"');
  expect(xml).toContain('r:id="link2"');
  expect(xml).toContain('<p:cNvPr id="10" name="Text"/>');
  const rels = await output
    .file('ppt/slides/_rels/slide1.xml.rels')!
    .async('string');
  expect(rels.match(/Id="link1"/g)).toHaveLength(1);
  expect(rels.match(/Id="img11"/g)).toHaveLength(1);
  expect(rels).toContain('Target="https://existing.example"');
});

it('lists commands in all presentation parts with default delimiters', async () => {
  const template = await presentation({
    'slides/slide1.xml': slide(shape(paragraph('{{ title }}'))),
    'notesSlides/notesSlide1.xml': slide(shape(paragraph('{{ INS notes }}'))),
  });
  expect(await listCommands(template)).toEqual([
    { raw: 'INS notes', type: 'INS', code: 'notes' },
    { raw: 'INS title', type: 'INS', code: 'title' },
  ]);
});

it('rejects PowerPoint HTML and no longer recognizes QUERY as a command', async () => {
  const template = await presentation({
    'slides/slide1.xml': slide(shape(paragraph('{{ HTML html }}'))),
  });
  await expect(
    createReport({ template, data: { html: '<p>Hello</p>' } })
  ).rejects.toThrow('HTML is only supported in Word');
  expect(BUILT_IN_COMMANDS).not.toContain('QUERY');
});

it('uses default braces in Word and calls async data resolvers without arguments', async () => {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  zip.file(
    'word/document.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{{ title }}</w:t></w:r></w:p></w:body></w:document>'
  );
  const template = await zip.generateAsync({ type: 'uint8array' });
  const resolver = jest.fn(async () => ({ title: 'Hello Word' }));
  const output = await JSZip.loadAsync(
    await createReport({ template, data: resolver })
  );
  expect(await output.file('word/document.xml')!.async('string')).toContain(
    'Hello Word'
  );
  expect(resolver).toHaveBeenCalledTimes(1);
  expect(resolver).toHaveBeenCalledWith();
  expect(await listCommands(template)).toEqual([
    { raw: 'INS title', type: 'INS', code: 'title' },
  ]);
});
