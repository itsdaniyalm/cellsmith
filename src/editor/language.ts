import type { languages } from 'monaco-editor';
import { ERROR_LITERALS } from '../core/tokenizer';
import { monaco } from './monaco';

export const LANGUAGE_ID = 'excel-formula';
export const THEME_DARK = 'formula-dark';
export const THEME_LIGHT = 'formula-light';

const errorLiteral = new RegExp('#(?:' + ERROR_LITERALS.map((e) => e.slice(1).replace(/[/?!.]/g, '\\$&')).join('|') + ')');

/**
 * Syntax highlighting. This is a lexical pass only; the semantic work (unknown functions,
 * arity, scoping) is done by the analyzer and surfaced as diagnostics and hovers.
 */
const monarch: languages.IMonarchLanguage = {
  defaultToken: '',
  brackets: [
    { open: '(', close: ')', token: 'delimiter.parenthesis' },
    { open: '{', close: '}', token: 'delimiter.curly' },
    { open: '[', close: ']', token: 'delimiter.square' },
  ],
  tokenizer: {
    root: [
      [/\s+/, 'white'],
      [/"/, 'string', '@string'],
      [errorLiteral, 'error.literal'],
      // A name immediately followed by `(` is a function call
      [/[A-Za-z_À-￿][\w.À-￿]*(?=\()/, 'function'],
      [/(?:TRUE|FALSE)(?![\w.(])/i, 'keyword.constant'],
      // Sheet prefixes: 'My Sheet'!  Sheet1!  Sheet1:Sheet3!  [Book.xlsx]Sheet1!
      [/(?:\[[^\]]*\])?(?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w.À-￿]*)(?::[A-Za-z_][\w.]*)?!/, 'type.sheet'],
      // Whole column / row ranges
      [/\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w(\[])/, 'variable.range'],
      // Cell references
      [/\$?[A-Za-z]{1,3}\$?\d+(?![\w.(\[])/, 'variable.cell'],
      [/(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/, 'number'],
      // Table references: Table1[...]  and bare [@Col]
      [/[A-Za-z_À-￿][\w.À-￿]*(?=\[)/, 'type.table'],
      [/\[/, { token: 'type.table', next: '@tableSpec' }],
      [/[A-Za-z_\\À-￿][\w.\\À-￿]*/, 'identifier'],
      [/<>|<=|>=|[+\-*/^&=<>%@#]/, 'operator'],
      [/[()]/, '@brackets'],
      [/[{}]/, '@brackets'],
      [/[,;]/, 'delimiter'],
      [/:/, 'delimiter.range'],
    ],
    string: [
      [/[^"]+/, 'string'],
      [/""/, 'string.escape'],
      [/"/, 'string', '@pop'],
    ],
    tableSpec: [
      [/'./, 'type.table'], // escaped character such as '[ or '#
      [/\[/, { token: 'type.table', next: '@push' }],
      [/\]/, { token: 'type.table', next: '@pop' }],
      [/[^\[\]']+/, 'type.table'],
    ],
  },
};

const configuration: languages.LanguageConfiguration = {
  brackets: [
    ['(', ')'],
    ['{', '}'],
    ['[', ']'],
  ],
  autoClosingPairs: [
    { open: '(', close: ')', notIn: ['string'] },
    { open: '{', close: '}', notIn: ['string'] },
    { open: '[', close: ']', notIn: ['string'] },
    { open: '"', close: '"', notIn: ['string'] },
  ],
  surroundingPairs: [
    { open: '(', close: ')' },
    { open: '{', close: '}' },
    { open: '[', close: ']' },
    { open: '"', close: '"' },
  ],
  indentationRules: {
    increaseIndentPattern: /[({[]\s*$/,
    decreaseIndentPattern: /^\s*[)}\]]/,
  },
  onEnterRules: [
    {
      // `(|)` + Enter  ->  indented blank line between the brackets
      beforeText: /[({[]\s*$/,
      afterText: /^\s*[)}\]]/,
      action: { indentAction: monaco.languages.IndentAction.IndentOutdent },
    },
    {
      beforeText: /[({[]\s*$/,
      action: { indentAction: monaco.languages.IndentAction.Indent },
    },
  ],
  wordPattern: /[A-Za-z_\\][\w.\\]*/,
};

const dark: Parameters<typeof monaco.editor.defineTheme>[1] = {
  base: 'vs-dark',
  inherit: true,
  rules: [
    { token: 'function', foreground: 'DCDCAA' },
    { token: 'number', foreground: 'B5CEA8' },
    { token: 'string', foreground: 'CE9178' },
    { token: 'string.escape', foreground: 'D7BA7D' },
    { token: 'operator', foreground: 'C8C8C8' },
    { token: 'variable.cell', foreground: '9CDCFE' },
    { token: 'variable.range', foreground: '9CDCFE' },
    { token: 'identifier', foreground: '4FC1FF' },
    { token: 'type.sheet', foreground: '4EC9B0' },
    { token: 'type.table', foreground: '4EC9B0' },
    { token: 'keyword.constant', foreground: '569CD6' },
    { token: 'error.literal', foreground: 'F48771' },
    { token: 'delimiter', foreground: 'C8C8C8' },
    { token: 'delimiter.range', foreground: 'C8C8C8' },
  ],
  colors: { 'editor.background': '#1f1f1f' },
};

const light: Parameters<typeof monaco.editor.defineTheme>[1] = {
  base: 'vs',
  inherit: true,
  rules: [
    { token: 'function', foreground: '795E26' },
    { token: 'number', foreground: '098658' },
    { token: 'string', foreground: 'A31515' },
    { token: 'string.escape', foreground: 'EE0000' },
    { token: 'operator', foreground: '333333' },
    { token: 'variable.cell', foreground: '0070C1' },
    { token: 'variable.range', foreground: '0070C1' },
    { token: 'identifier', foreground: '001080' },
    { token: 'type.sheet', foreground: '267F99' },
    { token: 'type.table', foreground: '267F99' },
    { token: 'keyword.constant', foreground: '0000FF' },
    { token: 'error.literal', foreground: 'CD3131' },
    { token: 'delimiter', foreground: '333333' },
    { token: 'delimiter.range', foreground: '333333' },
  ],
  colors: { 'editor.background': '#ffffff' },
};

let registered = false;
export function registerFormulaLanguage(): void {
  if (registered) return;
  registered = true;
  monaco.languages.register({ id: LANGUAGE_ID, aliases: ['Excel Formula'] });
  monaco.languages.setMonarchTokensProvider(LANGUAGE_ID, monarch);
  monaco.languages.setLanguageConfiguration(LANGUAGE_ID, configuration);
  monaco.editor.defineTheme(THEME_DARK, dark);
  monaco.editor.defineTheme(THEME_LIGHT, light);
}
