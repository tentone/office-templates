import JSZip from 'jszip';
import { createReport } from '../index';
import { parseXml } from '../xml';
import { Node } from '../types';
import { customRunJs } from './customRunJs';

async function makeTemplate(format: string, text: string, inTable = false) {
  const zip = new JSZip();
  const word = format === 'docx';
  const prefix = word ? 'w' : 'a';
  const paragraph = text
    .split('\n')
    .map(
      line =>
        `<${prefix}:p><${prefix}:r><${prefix}:t>${line}</${prefix}:t></${prefix}:r></${prefix}:p>`
    )
    .join('');
  const content = inTable
    ? word
      ? `<w:tbl><w:tr><w:tc>${paragraph}</w:tc></w:tr></w:tbl>`
      : `<p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="360000"/></a:tblGrid><a:tr h="360000"><a:tc><a:txBody><a:bodyPr/>${paragraph}</a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
    : word
    ? paragraph
    : `<p:sp><p:txBody><a:bodyPr/>${paragraph}</p:txBody></p:sp>`;
  const part = word ? 'word/document.xml' : 'ppt/slides/slide1.xml';
  zip.file(
    '[Content_Types].xml',
    `<Types><Override PartName="/${
      word ? part : 'ppt/presentation.xml'
    }" ContentType="application/vnd.openxmlformats-officedocument.${
      word ? 'wordprocessingml.document' : 'presentationml.presentation'
    }.main+xml"/></Types>`
  );
  if (word) {
    zip.file(
      part,
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${content}</w:body></w:document>`
    );
  } else {
    zip.file(
      'ppt/presentation.xml',
      '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
    );
    zip.file(
      part,
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>${content}</p:spTree></p:cSld></p:sld>`
    );
  }
  return { template: await zip.generateAsync({ type: 'uint8array' }), part };
}

function textContent(node: Node): string {
  if (node._fTextNode) return '';
  if (node._tag === 'w:t' || node._tag === 'a:t') {
    return node._children
      .map(child => (child._fTextNode ? child._text : ''))
      .join('');
  }
  return node._children.map(textContent).join('');
}

describe.each(['docx', 'pptx'])('%s inline nesting', format => {
  describe.each(['default', 'custom runJs'])('%s executor', executor => {
    const runJs = executor === 'custom runJs' ? customRunJs : undefined;
    async function render(text: string, data: object = {}, inTable = false) {
      const { template, part } = await makeTemplate(format, text, inTable);
      const output = await JSZip.loadAsync(
        await createReport({ template, data, runJs })
      );
      const xml = await output.file(part)!.async('string');
      expect(xml).not.toContain('{{');
      return textContent(await parseXml(xml));
    }

    it.each([
      [true, true, 'beforeABCafter'],
      [true, false, 'beforeACafter'],
      [false, true, 'beforeafter'],
      [false, false, 'beforeafter'],
    ])(
      'handles nested conditions (outer=%s, inner=%s)',
      async (outer, inner, expected) => {
        expect(
          await render(
            'before{{ IF outer }}A{{ IF inner }}B{{ END-IF }}C{{ END-IF }}after',
            { outer, inner }
          )
        ).toBe(expected);
      }
    );

    it('handles nested starts on the same line with endings in later paragraphs', async () => {
      expect(
        await render(
          'before{{ IF outer }}A{{ IF inner }}B\nbody\nC{{ END-IF }}D{{ END-IF }}after',
          { outer: true, inner: true }
        )
      ).toBe('beforeABbodyCDafter');
      expect(
        await render(
          'before{{ IF outer }}A{{ IF inner }}B\nbody\nC{{ END-IF }}D{{ END-IF }}after',
          { outer: true, inner: false }
        )
      ).toBe('beforeADafter');
      expect(
        await render(
          'before{{ FOR item IN items }}[{{ IF true }}{{ $item }}\n{{ END-IF }}]{{ END-FOR item }}after',
          { items: ['A', 'B'] }
        )
      ).toBe('before[A][B]after');
    });

    it('handles nested conditions inside a loop over multiple items', async () => {
      const text =
        'before{{ FOR item IN items }}{{ IF $item.show }}[{{ $item.name }}{{ IF $item.detail }}!{{ END-IF }}]{{ END-IF }}{{ END-FOR item }}after';
      expect(
        await render(text, {
          items: [
            { name: 'A', show: true, detail: true },
            { name: 'B', show: false, detail: true },
            { name: 'C', show: true, detail: false },
          ],
        })
      ).toBe('before[A!][C]after');
    });

    it('uses the FOR index in IF conditions and content inside the IF', async () => {
      expect(
        await render(
          '{{ FOR item IN items }}{{ IF $idx === 1 }}{{ $item }}:{{ $idx }}{{ END-IF }}{{ END-FOR item }}',
          { items: ['first', 'second', 'third'] }
        )
      ).toBe('second:1');
    });

    it('preserves the FOR index through multiple nested conditions', async () => {
      expect(
        await render(
          '{{ FOR item IN items }}{{ IF true }}{{ IF $idx === 2 }}{{ IF $idx > 0 }}{{ $item }}:{{ $idx }}{{ END-IF }}{{ END-IF }}{{ END-IF }}{{ END-FOR item }}',
          { items: ['first', 'second', 'third'] }
        )
      ).toBe('third:2');
    });

    it('uses the innermost FOR index inside IF and restores the outer index afterward', async () => {
      expect(
        await render(
          '{{ FOR group IN groups }}G{{ $idx }}[{{ IF true }}{{ FOR item IN $group.items }}{{ IF true }}{{ IF $idx === 1 }}{{ $item }}:{{ $idx }}{{ END-IF }}{{ END-IF }}{{ END-FOR item }}{{ END-IF }}]{{ $idx }};{{ END-FOR group }}',
          { groups: [{ items: ['A0', 'A1', 'A2'] }, { items: ['B0', 'B1'] }] }
        )
      ).toBe('G0[A1:1]0;G1[B1:1]1;');
    });

    it('handles nested loops with an empty inner collection', async () => {
      const text =
        'before{{ FOR group IN groups }}[{{ $group.name }}:{{ FOR value IN $group.values }}{{ $value }}{{ END-FOR value }}]{{ END-FOR group }}after';
      expect(
        await render(text, {
          groups: [
            { name: 'A', values: [1, 2] },
            { name: 'B', values: [] },
            { name: 'C', values: [3] },
          ],
        })
      ).toBe('before[A:12][B:][C:3]after');
    });

    it('handles a loop nested inside a condition', async () => {
      expect(
        await render(
          'before{{ IF show }}{{ FOR item IN items }}{{ $item }}{{ END-FOR item }}{{ END-IF }}after',
          { show: true, items: [1, 2] }
        )
      ).toBe('before12after');
    });

    it('skips nested expressions inside false conditions and empty loops', async () => {
      expect(
        await render(
          'before{{ IF false }}{{ IF missing.value }}hidden{{ END-IF }}{{ FOR item IN missing.items }}{{ $item }}{{ END-FOR item }}{{ END-IF }}after'
        )
      ).toBe('beforeafter');
      expect(
        await render(
          'before{{ FOR item IN items }}{{ IF missing.value }}hidden{{ END-IF }}{{ FOR inner IN missing.items }}{{ $inner }}{{ END-FOR inner }}{{ END-FOR item }}after',
          { items: [] }
        )
      ).toBe('beforeafter');
    });

    it('supports nested blocks in a table cell on the same row', async () => {
      expect(
        await render(
          'before{{ IF show }}{{ FOR item IN items }}{{ IF $item }}{{ $item }}{{ END-IF }}{{ END-FOR item }}{{ END-IF }}after',
          { show: true, items: [1, 0, 2] },
          true
        )
      ).toBe('before12after');
    });

    it('keeps adjacent blocks independent after closing nested conditions', async () => {
      expect(
        await render(
          '{{ IF true }}{{ IF true }}A{{ END-IF }}{{ END-IF }}{{ IF false }}B{{ END-IF }}{{ FOR item IN items }}{{ IF true }}{{ $item }}{{ END-IF }}{{ END-FOR item }}',
          { items: [1, 2] }
        )
      ).toBe('A12');
    });
  });
});
