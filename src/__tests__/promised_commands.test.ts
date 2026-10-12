import JSZip from 'jszip';
import {
  createReport,
  CommandExecutionError,
  NullishCommandResultError,
} from '../index';
import { customRunJs } from './customRunJs';

async function template(format: string, commands: string[]) {
  const zip = new JSZip();
  const namespaces =
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  if (format === 'docx') {
    zip.file(
      '[Content_Types].xml',
      '<Types><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    );
    zip.file(
      'word/document.xml',
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ${namespaces}><w:body>${commands
        .map(cmd => `<w:p><w:r><w:t>{{ ${cmd} }}</w:t></w:r></w:p>`)
        .join('')}</w:body></w:document>`
    );
  } else {
    zip.file(
      '[Content_Types].xml',
      '<Types><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>'
    );
    zip.file(
      'ppt/presentation.xml',
      '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
    );
    zip.file(
      'ppt/slides/slide1.xml',
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ${namespaces}><p:cSld><p:spTree>${commands
        .map(
          (cmd, i) =>
            `<p:sp><p:nvSpPr><p:cNvPr id="${
              i + 1
            }" name="Command"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{{ ${cmd} }}</a:t></a:r></a:p></p:txBody></p:sp>`
        )
        .join('')}</p:spTree></p:cSld></p:sld>`
    );
  }
  return zip.generateAsync({ type: 'uint8array' });
}

describe.each(['docx', 'pptx'])('%s promise commands', format => {
  describe.each(['default', 'custom runJs'])('%s execution', execution => {
    const runJs = execution === 'custom runJs' ? customRunJs : undefined;
    const part =
      format === 'docx' ? 'word/document.xml' : 'ppt/slides/slide1.xml';
    const image = {
      data: Buffer.from('resolved image bytes'),
      extension: '.png',
      width: 2,
      height: 3,
    };

    it('resolves INS and IMAGE values supplied as promises in data', async () => {
      const output = await JSZip.loadAsync(
        await createReport({
          template: await template(format, ['INS text', 'IMAGE image']),
          data: {
            text: Promise.resolve('Resolved & inserted'),
            image: Promise.resolve(image),
          },
          runJs,
        })
      );
      const xml = await output.file(part)!.async('string');
      expect(xml).toContain('Resolved &amp; inserted');
      expect(xml).not.toContain('[object Promise]');
      const media = Object.keys(output.files).find(
        path => path.includes('/media/') && path.endsWith('.png')
      )!;
      expect(await output.file(media)!.async('nodebuffer')).toEqual(image.data);
    });

    it('waits for asynchronous functions without await in template commands', async () => {
      const calls: string[] = [];
      const deferred = <T>(value: T) =>
        new Promise<T>(resolve => setImmediate(() => resolve(value)));
      const output = await JSZip.loadAsync(
        await createReport({
          template: await template(format, [
            'INS loadText()',
            'IMAGE loadImage()',
          ]),
          additionalJsContext: {
            loadText: async () => {
              const value = await deferred('Async text');
              calls.push('text');
              return value;
            },
            loadImage: async () => {
              expect(calls).toEqual(['text']);
              const value = await deferred(image);
              calls.push('image');
              return value;
            },
          },
          runJs,
        })
      );
      expect(await output.file(part)!.async('string')).toContain('Async text');
      expect(calls).toEqual(['text', 'image']);
      const media = Object.keys(output.files).find(
        path => path.includes('/media/') && path.endsWith('.png')
      )!;
      expect(await output.file(media)!.async('nodebuffer')).toEqual(image.data);
    });

    it.each(['INS', 'IMAGE'])(
      'reports rejected %s promises as execution errors',
      async command => {
        const failure = new Error('async lookup failed');
        await expect(
          createReport({
            template: await template(format, [`${command} load()`]),
            additionalJsContext: {
              load: async () => {
                throw failure;
              },
            },
            runJs,
          })
        ).rejects.toMatchObject({
          name: new CommandExecutionError(failure, 'load()').name,
          err: failure,
          command: 'load()',
        });
      }
    );

    it.each(['INS', 'IMAGE'])(
      'checks resolved %s values with rejectNullish',
      async command => {
        await expect(
          createReport({
            template: await template(format, [`${command} load()`]),
            additionalJsContext: { load: async () => null },
            runJs,
            rejectNullish: true,
          })
        ).rejects.toBeInstanceOf(NullishCommandResultError);
      }
    );
  });
});
