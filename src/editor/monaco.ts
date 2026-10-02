/**
 * Minimal Monaco build: the editor core plus only the features a formula editor uses.
 * `editor.main` would also register ~80 programming languages and the TS/CSS/HTML/JSON
 * language services, which is several megabytes we never touch.
 *
 * These are deep imports into monaco-editor's ESM sources, so the dependency is pinned to an
 * exact version in package.json. When upgrading, diff against `monaco-editor/esm/vs/editor/editor.main.js`.
 */
import * as api from '@monaco/editor/editor.api.js';

import '@monaco/editor/contrib/anchorSelect/browser/anchorSelect.js';
import '@monaco/editor/contrib/bracketMatching/browser/bracketMatching.js';
import '@monaco/editor/contrib/caretOperations/browser/caretOperations.js';
import '@monaco/editor/contrib/caretOperations/browser/transpose.js';
import '@monaco/editor/contrib/clipboard/browser/clipboard.js';
import '@monaco/editor/contrib/codeAction/browser/codeActionContributions.js';
import '@monaco/editor/browser/widget/codeEditor/codeEditorWidget.js';
import '@monaco/editor/contrib/comment/browser/comment.js';
import '@monaco/editor/contrib/contextmenu/browser/contextmenu.js';
import '@monaco/editor/contrib/cursorUndo/browser/cursorUndo.js';
import '@monaco/editor/contrib/find/browser/findController.js';
import '@monaco/editor/contrib/folding/browser/folding.js';
import '@monaco/editor/contrib/format/browser/formatActions.js';
import '@monaco/editor/contrib/gotoError/browser/gotoError.js';
import '@monaco/editor/contrib/gotoError/browser/markerSelectionStatus.js';
import '@monaco/editor/contrib/gotoSymbol/browser/goToCommands.js';
import '@monaco/editor/contrib/gotoSymbol/browser/link/goToDefinitionAtPosition.js';
import '@monaco/editor/contrib/hover/browser/hoverContribution.js';
import '@monaco/editor/contrib/indentation/browser/indentation.js';
import '@monaco/editor/contrib/linesOperations/browser/linesOperations.js';
import '@monaco/editor/contrib/multicursor/browser/multicursor.js';
import '@monaco/editor/contrib/parameterHints/browser/parameterHints.js';
import '@monaco/editor/contrib/placeholderText/browser/placeholderText.contribution.js';
import '@monaco/editor/contrib/rename/browser/rename.js';
import '@monaco/editor/contrib/smartSelect/browser/smartSelect.js';
import '@monaco/editor/contrib/snippet/browser/snippetController2.js';
import '@monaco/editor/contrib/suggest/browser/suggestController.js';
import '@monaco/editor/contrib/tokenization/browser/tokenization.js';
import '@monaco/editor/contrib/wordHighlighter/browser/wordHighlighter.js';
import '@monaco/editor/contrib/wordOperations/browser/wordOperations.js';
import '@monaco/editor/browser/coreCommands.js';
import '@monaco/editor/standalone/browser/quickAccess/standaloneCommandsQuickAccess.js';
import '@monaco/editor/standalone/browser/quickAccess/standaloneGotoLineQuickAccess.js';
import '@monaco/editor/common/standaloneStrings.js';
import '@monaco/base/browser/ui/codicons/codicon/codicon.css';
import '@monaco/base/browser/ui/codicons/codicon/codicon-modifiers.css';

import EditorWorker from './editor.worker?worker';

// The only language we register is our own (tokenized on the main thread), so the generic editor
// worker is all Monaco ever asks for.
(self as unknown as { MonacoEnvironment: unknown }).MonacoEnvironment = {
  getWorker: () => new EditorWorker(),
};

export const monaco = api as unknown as typeof import('monaco-editor');
export type Monaco = typeof monaco;
