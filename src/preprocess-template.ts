import {
  insertTextSiblingAfter,
  getNextSibling,
  normalizeTag,
} from './report-utils';
import { Node } from './types';
import { TemplateParseError } from './errors';

const styleCommand = /^\s*STYLE-TXT(?:\s|$)/i;

function canCloseStyleCommand(expression: string, delimiter: string) {
  // Skip quoted strings and comments while counting JavaScript object braces.
  const tokens =
    /"(?:\\[\s\S]|[^"\\])*"|'(?:\\[\s\S]|[^'\\])*'|`(?:\\[\s\S]|[^`\\])*`|\/\/[^\n]*(?=\n)|\/\*[\s\S]*?\*\/|([{}])|(['"`]|\/\/|\/\*)/g;
  let depth = 0;
  let token: RegExpExecArray | null;
  while ((token = tokens.exec(expression))) {
    if (token[2]) return false;
    if (token[1] === '{') depth += 1;
    if (token[1] === '}') depth -= 1;
  }
  return delimiter[0] !== '}' || depth <= 0;
}

// In-place
// In case of split commands (or even split delimiters), joins all the pieces
// at the starting node
const preprocessTemplate = (
  template: Node,
  delimiter: [string, string],
  preserveSpace: boolean,
  fixSmartQuotes = false
) => {
  let node: Node | null = template;
  let fCmd = false;
  let openNode = null;
  let idxDelimiter = 0;
  const placeholderCmd = `${delimiter[0]}CMD_NODE${delimiter[1]}`;

  while (node != null) {
    // Add `xml:space` attr `preserve` to `w:t` tags
    if (
      preserveSpace &&
      !node._fTextNode &&
      normalizeTag(node._tag) === 'w:t'
    ) {
      node._attrs['xml:space'] = 'preserve';
    }

    // Add a space if we reach a new `w:p` tag and there's an open node (hence, in a command)
    if (!node._fTextNode && normalizeTag(node._tag) === 'w:p' && openNode) {
      openNode._text += ' ';
    }

    // Process text nodes inside `w:t` tags
    if (
      node._fTextNode &&
      node._parent &&
      !node._parent._fTextNode &&
      normalizeTag(node._parent._tag) === 'w:t'
    ) {
      if (openNode == null) openNode = node;
      const textIn = node._text;
      node._text = '';
      for (let i = 0; i < textIn.length; i++) {
        const c = textIn[i];

        // What's the current expected delimiter
        const currentDelimiter = fCmd ? delimiter[1] : delimiter[0];

        // Matches the expected delimiter character
        if (c === currentDelimiter[idxDelimiter]) {
          const command = fCmd ? openNode._text.slice(delimiter[0].length) : '';
          const isStyleCommand = styleCommand.test(command);
          if (isStyleCommand && idxDelimiter === 0) {
            const expression = fixSmartQuotes
              ? command
                  .replace(/[\u201C\u201D\u201E]/g, '"')
                  .replace(/[\u2018\u2019\u201A]/g, "'")
              : command;
            if (!canCloseStyleCommand(expression, currentDelimiter)) {
              openNode._text += c;
              continue;
            }
          }
          idxDelimiter += 1;

          // Finished matching delimiter? Then toggle `fCmd`,
          // add a new `w:t` + text node (either before or after the delimiter),
          // depending on the case
          if (idxDelimiter === currentDelimiter.length) {
            fCmd = !fCmd;
            const fNodesMatch = node === openNode;
            if (fCmd && openNode._text.length) {
              openNode = insertTextSiblingAfter(openNode);
              if (fNodesMatch) node = openNode;
            }
            openNode._text += currentDelimiter;
            if (!fCmd && isStyleCommand) openNode._command = command;
            if (!fCmd && i < textIn.length - 1) {
              openNode = insertTextSiblingAfter(openNode);
              if (fNodesMatch) node = openNode;
            }
            idxDelimiter = 0;
            if (!fCmd) openNode = node; // may switch open node to the current one
          }

          // Doesn't match the delimiter, but we had some partial match
        } else if (idxDelimiter) {
          openNode._text += currentDelimiter.slice(0, idxDelimiter);
          idxDelimiter = 0;
          if (!fCmd) openNode = node;
          openNode._text += c;

          // General case
        } else {
          openNode._text += c;
        }
      }

      // Close the text node if nothing's pending
      if (!fCmd && !idxDelimiter) openNode = null;

      // If text was present but not any more, add a placeholder, so that this node
      // will be purged during report generation
      if (textIn.length && !node._text.length) node._text = placeholderCmd;
    }

    // Find next node to process
    if (node._children.length) node = node._children[0];
    else {
      let fFound = false;
      while (node._parent != null) {
        const nodeParent: Node = node._parent;
        const nextSibling = getNextSibling(node);
        if (nextSibling) {
          fFound = true;
          node = nextSibling;
          break;
        }
        node = nodeParent;
      }
      if (!fFound) node = null;
    }
  }
  if (
    fCmd &&
    openNode &&
    styleCommand.test(openNode._text.slice(delimiter[0].length))
  ) {
    throw new TemplateParseError('Unclosed STYLE-TXT command delimiter');
  }
  return template;
};

// ==========================================
// Public API
// ==========================================
export default preprocessTemplate;
