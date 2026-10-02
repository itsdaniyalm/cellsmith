export interface Settings {
  /** Preferred line width for formatting; 'auto' fits the task pane. */
  width: 'auto' | number;
  indent: 2 | 4;
  uppercaseFunctions: boolean;
  /** Expand LET onto several lines once it has this many bindings. */
  expandLetAt: number;
  /** Pretty-print single-line formulas when a cell is selected. */
  formatOnLoad: boolean;
  /** Write edits to the cell automatically once the formula is valid. */
  autoApply: boolean;
  /** 'as-typed' keeps your layout in the cell; 'minified' strips whitespace on write. */
  writeMode: 'as-typed' | 'minified';
}

export const DEFAULT_SETTINGS: Settings = {
  width: 'auto',
  indent: 4,
  uppercaseFunctions: true,
  expandLetAt: 2,
  formatOnLoad: true,
  autoApply: true,
  writeMode: 'as-typed',
};

const KEY = 'cellsmith.settings.v1';

/** Storage can be unavailable (private windows, blocked site data), so every access is guarded. */
export function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) return { ...DEFAULT_SETTINGS, ...(JSON.parse(raw) as Partial<Settings>) };
  } catch {
    /* fall through to defaults */
  }
  return { ...DEFAULT_SETTINGS };
}

export function saveSettings(s: Settings): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(s));
  } catch {
    /* not persisted; the session still works */
  }
}
