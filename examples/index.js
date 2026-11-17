const fs = require('fs').promises;
const path = require('path');
const vm = require('vm');
const qrCode = require('qrcode');
const {
  createReport,
  getMetadata,
  listCommands,
  NullishCommandResultError,
} = require('..');

const outputDir = path.join(__dirname, 'output');
const repositoryUrl = 'https://github.com/tentone/office-templates';

async function loadReportData() {
  return {
    title: 'Community garden update',
    summary: Promise.resolve(
      'Spring planting is complete.\nThe next workday is Saturday.'
    ),
    beds: [
      {
        name: 'Herb bed',
        crops: ['Basil', 'Mint'],
        ready: true,
        plots: [3, 2],
      },
      {
        name: 'Salad bed',
        crops: ['Lettuce', 'Chard'],
        ready: false,
        plots: [4, 1],
      },
    ],
    plotSeasons: ['Spring', 'Summer'],
    statusStyle: { color: '#226143', bold: true },
    detailStyle: { italic: true },
    coordinator: null,
    repositoryLink: { url: repositoryUrl, label: 'Office Templates on GitHub' },
    scheduleHtml: `<!DOCTYPE html>
      <html lang="en">
        <head><meta charset="UTF-8"></head>
        <body><p><strong>Next review</strong> is scheduled for Friday.</p></body>
      </html>`,
    // Write a Word tab as a sibling of the surrounding text elements.
    literalXmlText:
      'Plot map||</w:t><w:tab/><w:t xml:space="preserve">||Garden entrance',
  };
}

const additionalJsContext = {
  makeQrCode: async () => ({
    data: await qrCode.toBuffer(repositoryUrl, { type: 'png' }),
    extension: '.png',
    width: 2.6,
    height: 2.6,
    alt: 'QR code for Office Templates on GitHub',
  }),
};

function errorHandler(error) {
  if (error instanceof NullishCommandResultError) return 'To be confirmed';
  throw error;
}

// A custom Node.js executor for the bundled, trusted slide template.
function runJs({ sandbox }) {
  const result = vm.runInNewContext(sandbox.__code__, sandbox);
  return { modifiedSandbox: sandbox, result };
}

async function main() {
  const [wordTemplate, slidesTemplate] = await Promise.all([
    fs.readFile(path.join(__dirname, 'template.docx')),
    fs.readFile(path.join(__dirname, 'template.pptx')),
  ]);

  console.log('Word template metadata', await getMetadata(wordTemplate));
  console.table(await listCommands(slidesTemplate));

  const slidesData = await loadReportData();
  // Both templates use the default {{ and }} command delimiters.
  const [wordReport, slidesReport] = await Promise.all([
    createReport({
      template: wordTemplate,
      data: loadReportData,
      additionalJsContext,
      rejectNullish: true,
      errorHandler,
    }),
    createReport({
      template: slidesTemplate,
      data: slidesData,
      additionalJsContext,
      runJs,
    }),
  ]);

  await fs.mkdir(outputDir, { recursive: true });
  await fs.writeFile(path.join(outputDir, 'report.docx'), wordReport);
  await fs.writeFile(path.join(outputDir, 'report.pptx'), slidesReport);
  console.log(`Wrote report.docx and report.pptx to ${outputDir}`);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
