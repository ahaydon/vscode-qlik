import { randomUUID } from 'crypto';

export interface ScriptSection {
  /** Stable identifier — survives reordering */
  id: string;
  name: string;
  body: string;
}

const TAB_LINE = /^\/\/\/\$tab (.+)$/;

/**
 * Parse a Qlik load script into ordered sections.
 * Sections are delimited by `///$tab <name>` lines.
 */
export function parseScript(raw: string): ScriptSection[] {
  const lines = raw.split(/\r?\n/);
  const sections: ScriptSection[] = [];
  let currentName: string | null = null;
  let bodyLines: string[] = [];

  const flush = () => {
    if (currentName !== null) {
      // Trim trailing blank lines from body
      let end = bodyLines.length;
      while (end > 0 && bodyLines[end - 1].trim() === '') end--;
      sections.push({ id: randomUUID(), name: currentName, body: bodyLines.slice(0, end).join('\n') });
      bodyLines = [];
    }
  };

  for (const line of lines) {
    const m = line.match(TAB_LINE);
    if (m) {
      flush();
      currentName = m[1].trim();
    } else if (currentName !== null) {
      bodyLines.push(line);
    }
    // Lines before the first ///$tab are ignored (Qlik ignores them too)
  }
  flush();

  if (sections.length === 0) {
    // Script has no section markers — treat whole script as Main
    sections.push({ id: randomUUID(), name: 'Main', body: raw.trimEnd() });
  }

  return sections;
}

/**
 * Serialize ordered sections back into a single Qlik load script string.
 */
export function serializeScript(sections: ScriptSection[]): string {
  return sections
    .map(s => `///$tab ${s.name}\n${s.body}`)
    .join('\n');
}
