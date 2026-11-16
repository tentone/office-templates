import { InvalidCommandError } from './errors';
import { addChild, newNonTextNode, newTextNode } from './report-utils';
import type { Node, NonTextNode, TextNode } from './types';

const alignments = ['left', 'center', 'right', 'justify'] as const;
const directions = ['ltr', 'rtl'] as const;
const highlightColors = [
  'black',
  'blue',
  'cyan',
  'green',
  'magenta',
  'red',
  'yellow',
  'white',
  'darkBlue',
  'darkCyan',
  'darkGreen',
  'darkMagenta',
  'darkRed',
  'darkYellow',
  'darkGray',
  'lightGray',
];
const hexColor = /^#[0-9a-f]{6}$/i;

export type WordTextStyle = {
  color?: string;
  highlightColor?: string | null;
  alignment?: (typeof alignments)[number];
  direction?: (typeof directions)[number];
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  striketrough?: boolean;
  verticalAlign?: 'subscript' | 'superscript' | 'baseline';
  font?: string;
  fontSize?: number;
};

export function resolveTextStyle(
  value: unknown,
  inherited: WordTextStyle | undefined,
  command: string
): WordTextStyle | undefined {
  if (Object.prototype.toString.call(value) !== '[object Object]') {
    throw new InvalidCommandError('STYLE-TXT expects a style object', command);
  }
  const input = value as Record<string, unknown>;
  const style = { ...inherited };
  const invalid = (name: string) =>
    new InvalidCommandError(`Invalid STYLE-TXT ${name}`, command);

  const readBoolean = (name: string) => {
    const value = input[name];
    if (value !== undefined && typeof value !== 'boolean') throw invalid(name);
    return value;
  };
  for (const key of ['bold', 'italic', 'underline', 'striketrough'] as const) {
    const value =
      key === 'striketrough' && input[key] === undefined
        ? readBoolean('strikethrough')
        : readBoolean(key);
    if (value !== undefined) style[key] = value;
  }

  const subscript = readBoolean('subscript');
  const superscript = readBoolean('superscript');
  if (subscript && superscript) {
    throw invalid('object (subscript and superscript cannot both be true)');
  }
  if (subscript) style.verticalAlign = 'subscript';
  else if (superscript) style.verticalAlign = 'superscript';
  else if (
    (subscript === false && style.verticalAlign !== 'superscript') ||
    (superscript === false && style.verticalAlign !== 'subscript')
  ) {
    style.verticalAlign = 'baseline';
  }

  const { color, highlightColor, alignment, direction, font, fontSize } = input;
  if (color !== undefined) {
    if (typeof color !== 'string' || !hexColor.test(color))
      throw invalid('color');
    style.color = color.slice(1).toUpperCase();
  }
  if (highlightColor === null) style.highlightColor = null;
  else if (highlightColor !== undefined) {
    if (typeof highlightColor !== 'string') throw invalid('highlightColor');
    const named = highlightColors.find(
      color => color.toLowerCase() === highlightColor.toLowerCase()
    );
    if (!named && !hexColor.test(highlightColor))
      throw invalid('highlightColor');
    style.highlightColor = named ?? highlightColor.toUpperCase();
  }
  if (alignment !== undefined) {
    const match = alignments.find(value => value === alignment);
    if (!match) throw invalid('alignment');
    style.alignment = match;
  }
  if (direction !== undefined) {
    const match = directions.find(value => value === direction);
    if (!match) throw invalid('direction');
    style.direction = match;
  }
  if (font !== undefined) {
    if (typeof font !== 'string' || !font.trim()) throw invalid('font');
    style.font = font;
  }
  if (fontSize !== undefined) {
    if (
      typeof fontSize !== 'number' ||
      !Number.isFinite(fontSize) ||
      fontSize < 1 ||
      fontSize > 1638
    ) {
      throw invalid('fontSize (expected 1 to 1638 points)');
    }
    style.fontSize = Math.round(fontSize * 2);
  }
  return Object.keys(style).length ? style : undefined;
}

const isElement = <Tag extends string>(
  node: Node,
  tag: Tag
): node is NonTextNode & { _tag: Tag } => !node._fTextNode && node._tag === tag;

function copyNode(node: Node): Node {
  return node._fTextNode
    ? newTextNode(node._text)
    : newNonTextNode(node._tag, node._attrs, node._children.map(copyNode));
}

function setProperty(
  parent: NonTextNode,
  property: NonTextNode,
  before: readonly string[]
) {
  const children = parent._children;
  const existing = children.findIndex(child => isElement(child, property._tag));
  const following = children.findIndex(
    child => !child._fTextNode && before.includes(child._tag)
  );
  property._parent = parent;
  if (existing >= 0) children.splice(existing, 1, property);
  else
    children.splice(following >= 0 ? following : children.length, 0, property);
}

const property = (name: string, value: string) =>
  newNonTextNode(`w:${name}`, { 'w:val': value });

const toggle = (value: boolean | undefined) =>
  value === undefined ? undefined : value ? '1' : '0';

function applyRunProperties(props: NonTextNode, style: WordTextStyle) {
  const settings: Record<string, string | undefined> = {
    color: style.color,
    b: toggle(style.bold),
    bCs: toggle(style.bold),
    i: toggle(style.italic),
    iCs: toggle(style.italic),
    u:
      style.underline === undefined
        ? undefined
        : style.underline
        ? 'single'
        : 'none',
    strike: toggle(style.striketrough),
    dstrike: style.striketrough === undefined ? undefined : '0',
    vertAlign: style.verticalAlign,
    rtl:
      style.direction === undefined
        ? undefined
        : toggle(style.direction === 'rtl'),
    sz: style.fontSize?.toString(),
    szCs: style.fontSize?.toString(),
  };
  const changes: NonTextNode[] = [];
  for (const [name, value] of Object.entries(settings)) {
    if (value !== undefined) changes.push(property(name, value));
  }
  if (style.font !== undefined) {
    changes.push(
      newNonTextNode('w:rFonts', {
        'w:ascii': style.font,
        'w:hAnsi': style.font,
        'w:eastAsia': style.font,
        'w:cs': style.font,
      })
    );
  }
  const { highlightColor } = style;
  if (highlightColor !== undefined) {
    // Word's native highlight has a fixed palette. Use shading for RGB colors.
    const rgb = highlightColor?.startsWith('#')
      ? highlightColor.slice(1)
      : undefined;
    changes.push(
      property('highlight', rgb ? 'none' : highlightColor ?? 'none')
    );
    if (rgb || highlightColor === null) {
      changes.push(
        newNonTextNode('w:shd', {
          'w:val': rgb ? 'clear' : 'nil',
          'w:fill': rgb ?? 'auto',
        })
      );
    }
  }
  changes.forEach(change => setProperty(props, change, ['w:rPrChange']));
}

const afterJustification = [
  'w:textDirection',
  'w:textAlignment',
  'w:textboxTightWrap',
  'w:outlineLvl',
  'w:divId',
  'w:rPr',
  'w:sectPr',
  'w:pPrChange',
];

export function applyWordTextStyles(
  root: Node,
  styles: WeakMap<TextNode, WordTextStyle>
) {
  type ParagraphStyle = Pick<WordTextStyle, 'alignment'>;

  function splitRun(
    run: NonTextNode,
    paragraph?: ParagraphStyle
  ): NonTextNode[] {
    const children = run._children;
    if (
      !children.some(
        child =>
          isElement(child, 'w:t') &&
          child._children.some(text => text._fTextNode && styles.has(text))
      )
    ) {
      return [run];
    }

    const props = children.find(child => isElement(child, 'w:rPr'));
    // Keep the template properties for each new run, including unstyled text.
    const runs: NonTextNode[] = [];
    let current: NonTextNode | undefined;
    let previousStyle: WordTextStyle | undefined;
    function emit(child: Node, style?: WordTextStyle) {
      if (!current || style !== previousStyle) {
        current = newNonTextNode(run._tag, run._attrs);
        current._parent = run._parent;
        runs.push(current);
        previousStyle = style;
        const runProps = props
          ? newNonTextNode(
              props._tag,
              props._attrs,
              props._children.map(copyNode)
            )
          : newNonTextNode('w:rPr');
        if (style) applyRunProperties(runProps, style);
        if (props || runProps._children.length) addChild(current, runProps);
      }
      addChild(current, child);
      if (paragraph && style?.alignment) paragraph.alignment = style.alignment;
    }

    for (const child of children) {
      if (child === props || (child._fTextNode && !child._text.trim()))
        continue;
      if (!isElement(child, 'w:t')) emit(child);
      else {
        for (const text of child._children) {
          if (text._fTextNode && !text._text) continue;
          emit(
            newNonTextNode(child._tag, child._attrs, [text]),
            text._fTextNode ? styles.get(text) : undefined
          );
        }
      }
    }
    return runs;
  }

  function visit(node: Node, paragraph?: ParagraphStyle) {
    if (node._fTextNode || !node._children.length) return;
    const isParagraph = node._tag === 'w:p';
    if (isParagraph) paragraph = {};
    const children: Node[] = [];
    for (const child of node._children) {
      visit(child, paragraph);
      if (isElement(child, 'w:r')) children.push(...splitRun(child, paragraph));
      else children.push(child);
    }
    node._children = children;
    if (isParagraph && paragraph?.alignment) {
      let props: NonTextNode | undefined = node._children.find(child =>
        isElement(child, 'w:pPr')
      );
      if (!props) {
        props = newNonTextNode('w:pPr');
        props._parent = node;
        node._children.unshift(props);
      }
      const alignment =
        paragraph.alignment === 'justify' ? 'both' : paragraph.alignment;
      setProperty(props, property('jc', alignment), afterJustification);
    }
  }

  visit(root);
}
