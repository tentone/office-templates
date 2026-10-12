// ==========================================
// Office XML nodes

import { QualifiedAttribute } from 'sax';

// ==========================================
export type Node = TextNode | NonTextNode;
type BaseNode = {
  _parent?: Node;
  _children: Array<Node>;
  _ifName?: string;
};
export type TextNode = BaseNode & {
  _fTextNode: true;
  _text: string;
};
export type NonTextNode = BaseNode & {
  _fTextNode: false;
  _tag: string;

  // Simplified; only need this property
  _attrs: { [key: string]: QualifiedAttribute | string } & {
    Extension?: string;
    ContentType?: string;
    PartName?: string;
  };
};

// ==========================================
// Report creator
// ==========================================
export type ReportData = any;
export type DataResolver = () => ReportData | Promise<ReportData>;

type ErrorHandler = (e: Error, raw_code?: string) => any;

export type RunJSResult = {
  modifiedSandbox: SandBox;
  result: unknown;
};

export type RunJSFunc = (o: {
  sandbox: SandBox;
  ctx: Context;
}) => RunJSResult | Promise<RunJSResult>;

export type UserOptions = {
  /**
   * Docx or pptx file template as a Uint8Array (or e.g. ArrayBuffer or NodeJS Buffer).
   */
  template: Uint8Array;
  /**
   * Object of data to be injected or a (async) function that resolves to the data. The function is called once without arguments.
   */
  data?: ReportData | DataResolver;
  /**
   * Defines a custom command delimiter. Defaults to ['{{', '}}']. This can be a String e.g. '+++' or an Array of Strings with length 2: ['{', '}'] in which the first element serves as the start delimiter and the second as the end delimiter.
   */
  cmdDelimiter?: string | [string, string];
  /**
   * The delimiter that's used to indicate literal XML that should be inserted into the document XML tree as-is. Defaults to `||`.
   */
  literalXmlDelimiter?: string;
  /**
   * Handle linebreaks in result of commands as actual linebreaks (Default: true)
   */
  processLineBreaks?: boolean; // true by default
  /**
   * Override JavaScript execution, for example with a caller-provided sandbox.
   * Receives the command in sandbox.__code__ and the template context.
   * Return { modifiedSandbox, result }; both the return value and result may be promises.
   * Without this option, commands execute directly without isolation.
   */
  runJs?: RunJSFunc;
  /**
   * Add functions or other static data to this option to have access to it in your commands.
   *
   * ```js
   * additionalJsContext: {
   *   qrCode: url => {
   *     const dataUrl = createQrImage(url, { size: 500 });
   *     const data = dataUrl.slice('data:image/gif;base64,'.length);
   *     return { width: 6, height: 6, data, extension: '.gif' };
   *   },
   * }
   * ```
   */
  additionalJsContext?: Object;
  /**
   * Whether to fail on the first error encountered in the template. Defaults to true. Can be used to collect all errors in a template (e.g. misspelled commands) before failing.
   */
  failFast?: boolean;

  /**
   * When set to `true`, this setting ensures `createReport` throws a `NullishCommandResultError` when the result of an INS, HTML, IMAGE, or LINK command is `null` or `undefined`. This is useful as nullish return values usually indicate a mistake in the template or the invoking code. Defaults to `false`.
   */
  rejectNullish?: boolean;

  /**
   * Custom error handler to catch any errors that may occur evaluating commands in the template. The value returned from this handler will be inserted into the template instead.
   */
  errorHandler?: ErrorHandler;

  /**
   * MS Word usually autocorrects JS string literal quotes with unicode 'smart' quotes ('curly' quotes). E.g. 'aubergine' -> ‘aubergine’.
   * This causes an error when evaluating commands containing these smart quotes, as they are not valid JavaScript.
   * If you set fixSmartQuotes to 'true', these smart quotes will automatically get replaced with straight quotes (') before command evaluation.
   * Defaults to false.
   */
  fixSmartQuotes?: boolean;

  /**
   * Use the new way of injecting line breaks from command results (only applies when `processLineBreaks` is `true`)
   * which has better results in LibreOffice and Google Drive.
   * (Default: false)
   */
  processLineBreaksAsNewText?: boolean;

  /**
   * Maximum loop iterations allowed when walking through the template.
   * You can increase this to generate reports with large amount of FOR loop elements.
   * Tip: You can disable infinite loop protection by using the `Infinity` constant.
   * This may be useful if you implement a process timeout instead.
   * (Default: 1,000,000)
   */
  maximumWalkingDepth?: number;
  /**
   * Whether to indent the generated XML to make it more human-readable.
   * Tip: Set this to false if you want to minimize the size of the generated Office file.
   * (Default: true)
   */
  indentXml?: boolean;
  /**
   * Whether to preserve whitespace in the generated XML.*
   * Tip: Set this to false if you want to minimize the size of the generated Office file.
   * (Default: true)
   */
  preserveSpace?: boolean;
  /**
   * Compression level for the generated Office file.
   * Integer between 0 (no compression, fastest) and 9 (maximum compression, slowest).
   * (Default: 1)
   */
  compressionLevel?: number;
};

export type CreateReportOptions = {
  documentType?: 'docx' | 'pptx';
  cmdDelimiter: [string, string];
  literalXmlDelimiter: string;
  processLineBreaks: boolean;
  runJs?: RunJSFunc;
  additionalJsContext: Object;
  failFast: boolean;
  rejectNullish: boolean;
  errorHandler: ErrorHandler | null;
  fixSmartQuotes: boolean;
  processLineBreaksAsNewText: boolean;
  maximumWalkingDepth?: number;
  indentXml: boolean;
  preserveSpace: boolean;
  compressionLevel: number;
};

export type SandBox = {
  __code__: string | undefined;
  __result__: unknown | undefined;
  [k: string]: unknown;
};

export type Context = {
  gCntIf: number;
  gCntEndIf: number;
  level: number;
  fCmd: boolean;
  cmd: string;
  buffers: {
    'w:p': BufferStatus;
    'w:tr': BufferStatus;
    'w:tc': BufferStatus;
  };
  pendingImageNode?: { image: NonTextNode; caption?: NonTextNode[] };
  imageAndShapeIdIncrement: number;
  images: Images;
  relationshipIds: Set<string>;
  pendingLinkNode?: NonTextNode;
  linkId: number;
  links: Links;
  pendingHtmlNode?: TextNode | NonTextNode;
  htmlId: number;
  htmls: Htmls;
  vars: { [name: string]: VarValue };
  loops: Array<LoopStatus>;
  fJump: boolean;
  shorthands: { [shorthand: string]: string };
  options: CreateReportOptions;
  jsSandbox?: SandBox;
  textRunPropsNode?: NonTextNode;

  // To verfiy we don't have a nested if within the same p or tr tag
  pIfCheckMap: Map<Node, string>;
  trIfCheckMap: Map<Node, string>;
};

export type Images = { [id: string]: Image };
export const ImageExtensions = [
  '.png',
  '.gif',
  '.jpg',
  '.jpeg',
  '.svg',
] as const;
type ImageExtension = (typeof ImageExtensions)[number];
export type Image = {
  extension: ImageExtension;
  data: ArrayBuffer | string;
};
export type Links = { [id: string]: Link };
type Link = { url: string };
export type Htmls = { [id: string]: string };

type BufferStatus = {
  text: string;
  cmds: string;
  fInsertedText: boolean;
};

type VarValue = unknown;
export type LoopStatus = {
  refNode: Node;
  refNodeLevel: number;
  varName: string;
  loopOver: Array<VarValue>;
  idx: number;
  isIf?: boolean;
};

export type ImagePars = {
  /**
   * Desired width of the image in centimeters.
   */
  width: number;

  /**
   * Desired height of the image in centimeters.
   */
  height: number;

  /**
   * Either an ArrayBuffer or a base64 string with the image data.
   */
  data: ArrayBuffer | string;

  /**
   * Optional. When injecting an SVG image, a fallback non-SVG (png/jpg/gif, etc.) image can be provided. This thumbnail is used when SVG images are not supported (e.g. older versions of Word) or when the document is previewed by e.g. Windows Explorer. See usage example below.
   */
  thumbnail?: Image;

  /**
   * One of '.png', '.gif', '.jpg', '.jpeg', '.svg'.
   */
  extension: ImageExtension;

  /**
   * Optional alt text.
   */
  alt?: string;

  /**
   * Optional rotation in degrees, with positive angles moving clockwise.
   */
  rotation?: number;

  /**
   * Optional caption
   */
  caption?: string;
};

export type LinkPars = {
  url: string;
  label?: string;
};

export type CommandSummary = {
  raw: string;
  type: BuiltInCommand;
  code: string;
};

export type BuiltInCommand = (typeof BUILT_IN_COMMANDS)[number];

export const BUILT_IN_COMMANDS = [
  'CMD_NODE',
  'ALIAS',
  'FOR',
  'END-FOR',
  'IF',
  'END-IF',
  'INS',
  'EXEC',
  'IMAGE',
  'LINK',
  'HTML',
] as const;
