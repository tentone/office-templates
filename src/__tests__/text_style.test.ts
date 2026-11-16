import JSZip from 'jszip';
import { createReport, listCommands } from '../index';
import type { UserOptions } from '../types';
import { customRunJs } from './customRunJs';

async function word(body: string, header?: string) {
  const zip = new JSZip();
  const ns =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  zip.file(
    '[Content_Types].xml',
    '<Types><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  zip.file(
    'word/document.xml',
    `<w:document ${ns}><w:body>${body}</w:body></w:document>`
  );
  if (header) zip.file('word/header1.xml', `<w:hdr ${ns}>${header}</w:hdr>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const run = (text: string, properties = '') =>
  `<w:r>${properties}<w:t>${text}</w:t></w:r>`;
const paragraph = (text: string) => `<w:p>${run(text)}</w:p>`;

async function render(body: string, options: Partial<UserOptions> = {}) {
  return createReport(
    { template: await word(body), indentXml: false, ...options },
    'XML'
  );
}

function runs(xml: string) {
  const pattern = /<w:r>([\s\S]*?)<\/w:r>/g;
  const text = /<w:t[^>]*>([^<]+)<\/w:t>/;
  return Array.from(xml.matchAll(pattern), ([, content]) => [
    text.exec(content)?.[1],
    content,
  ]).filter(([value]) => value);
}

it('splits a single run at style boundaries and restores the enclosing and template styles', async () => {
  const content =
    'Before {{STYLE-TXT {bold:true, color:"#ff0000"}}}Red ' +
    '{{STYLE-TXT {bold:false, color: someColor}}}{{INS label}}{{STYLE-TXT-END}} red again' +
    '{{STYLE-TXT-END}} after';
  const props =
    '<w:rPr><w:color w:val="222222" w:themeColor="accent1"/><w:i/><w:lang w:val="en-US"/></w:rPr>';
  const result = await render(`<w:p>${run(content, props)}</w:p>`, {
    data: { someColor: '#00FF00', label: 'Green' },
    runJs: customRunJs,
  });
  const output = runs(result);
  expect(output.map(([text]) => text)).toEqual([
    'Before ',
    'Red ',
    'Green',
    ' red again',
    ' after',
  ]);
  expect(output[0][1]).toContain(props);
  expect(output[1][1]).toMatch(/w:color w:val="FF0000".*w:b w:val="1"/);
  expect(output[2][1]).toMatch(/w:color w:val="00FF00".*w:b w:val="0"/);
  expect(output[3][1]).toMatch(/w:color w:val="FF0000".*w:b w:val="1"/);
  expect(output[4][1]).toContain(props);
  for (const [, xml] of output)
    expect(xml).toContain('<w:lang w:val="en-US"/>');
});

it('applies run properties and paragraph alignment across split tags and paragraphs', async () => {
  const output = await render(
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/></w:pPr>${run(
      '{'
    )}${run(
      '{STYLE-TXT {alignment: "justify", direction: "rtl", highlightColor: "#13b3cc", font: "Arial", fontSize: 13.5, italic: true, underline: true, striketrough: true, subscript: true}'
    )}${run('}}First')}</w:p>` +
      paragraph(
        'Second {{STYLE-TXT {direction:"ltr", alignment:"center", highlightColor: null, superscript:true}}}Inner{{STYLE-TXT-END}}{{STYLE-TXT-END}} plain'
      ) +
      paragraph(
        '{{STYLE-TXT {alignment:"right", highlightColor:"darkBlue", underline:false, strikethrough:false, italic:false, superscript:false}}}Third{{STYLE-TXT-END}}'
      )
  );
  expect(output).toContain(
    '<w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="both"/><w:outlineLvl w:val="0"/></w:pPr>'
  );
  expect(output).toContain('<w:pPr><w:jc w:val="center"/></w:pPr>');
  expect(output).toContain('<w:pPr><w:jc w:val="right"/></w:pPr>');
  expect(output).not.toContain('STYLE-TXT');
  const formatted = Object.fromEntries(runs(output));
  for (const text of ['First', 'Second ']) {
    expect(formatted[text]).toContain('<w:rtl w:val="1"/>');
    expect(formatted[text]).toContain('<w:shd w:val="clear" w:fill="13B3CC"/>');
    expect(formatted[text]).toContain('<w:sz w:val="27"/><w:szCs w:val="27"/>');
    expect(formatted[text]).toContain(
      '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>'
    );
    expect(formatted[text]).toContain('<w:vertAlign w:val="subscript"/>');
  }
  expect(formatted.Inner).toContain('<w:rtl w:val="0"/>');
  expect(formatted.Inner).toContain(
    '<w:highlight w:val="none"/><w:shd w:val="nil" w:fill="auto"/>'
  );
  expect(formatted.Inner).toContain('<w:vertAlign w:val="superscript"/>');
  expect(formatted.Inner).toContain(
    '<w:i w:val="1"/><w:iCs w:val="1"/><w:u w:val="single"/><w:strike w:val="1"/>'
  );
  expect(formatted.Third).toContain(
    '<w:i w:val="0"/><w:iCs w:val="0"/><w:u w:val="none"/><w:strike w:val="0"/>'
  );
  expect(formatted.Third).toContain('<w:highlight w:val="darkBlue"/>');
  expect(formatted.Third).toContain('<w:vertAlign w:val="baseline"/>');
  expect(formatted[' plain']).not.toContain('<w:rPr>');
});

it('resolves styles for each loop iteration and skips styles in a false conditional', async () => {
  const output = await render(
    paragraph(
      '{{STYLE-TXT {bold:true}}}{{FOR item IN items}}{{STYLE-TXT {color: $item.color}}}' +
        '{{INS $item.label}}{{STYLE-TXT-END}}{{IF false}}{{STYLE-TXT missing}}Hidden{{STYLE-TXT-END}}' +
        '{{END-IF}}{{END-FOR item}}Tail{{STYLE-TXT-END}} Outside'
    ),
    {
      data: {
        items: [
          { label: 'A', color: '#FF0000' },
          { label: 'B', color: '#0000FF' },
        ],
      },
    }
  );
  const formatted = Object.fromEntries(runs(output));
  expect(Object.keys(formatted)).toEqual(['A', 'B', 'Tail', ' Outside']);
  expect(formatted.A).toContain('<w:color w:val="FF0000"/>');
  expect(formatted.B).toContain('<w:color w:val="0000FF"/>');
  expect(formatted.Tail).toContain('<w:b w:val="1"/>');
  expect(formatted.Tail).not.toContain('<w:color');
  expect(formatted[' Outside']).not.toContain('<w:rPr>');
});

it('lists both tags in a header and handles delimiters inside quoted style values', async () => {
  const code = '{font: "Name }} with braces", color: pick({value: "#FFFFFF"})}';
  const template = await word(
    paragraph('Body'),
    paragraph(`{{STYLE-TXT ${code}}}Header{{STYLE-TXT-END}}`)
  );
  expect(await listCommands(template)).toEqual([
    { raw: `STYLE-TXT ${code}`, type: 'STYLE-TXT', code },
    { raw: 'STYLE-TXT-END', type: 'STYLE-TXT-END', code: '' },
  ]);
  const output = await JSZip.loadAsync(
    await createReport({
      template,
      additionalJsContext: { pick: ({ value }: { value: string }) => value },
    })
  );
  const header = await output.file('word/header1.xml')!.async('string');
  expect(header).toContain('w:ascii="Name }} with braces"');
  expect(header).toContain('w:val="FFFFFF"');
  expect(header).toContain('>Header</w:t>');
});

it('handles custom braces and Word smart quotes in style expressions', async () => {
  const output = await render(
    paragraph(
      '{STYLE-TXT {font: “Font } Name”, color: “#FF0000”}}Text{STYLE-TXT-END}'
    ),
    { cmdDelimiter: ['{', '}'], fixSmartQuotes: true }
  );
  expect(output).toContain('w:ascii="Font } Name"');
  expect(output).toContain('<w:color w:val="FF0000"/>');
  expect(output).toContain('>Text</w:t>');
});

it.each([
  ['{{STYLE-TXT-END}}', 'Unexpected STYLE-TXT-END'],
  ['{{STYLE-TXT style}}Text', 'Missing STYLE-TXT-END'],
  ['{{STYLE-TXT false}}Text{{STYLE-TXT-END}}', 'expects a style object'],
  [
    '{{STYLE-TXT {fontSize: -2}}}Text{{STYLE-TXT-END}}',
    'Invalid STYLE-TXT fontSize',
  ],
])('reports invalid style commands in %s', async (content, error) => {
  await expect(
    render(paragraph(content), { data: { style: {} } })
  ).rejects.toThrow(error);
});
