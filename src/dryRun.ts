/**
 * Human-readable, copy-pasteable rendering of a RequestSpec for `--dry-run`.
 * Must never require a token or touch the network — RequestSpec is already
 * fully materialized by GraphClient#buildSpec by the time this runs.
 */

import { redactHeaders, redactObject } from './redact.js';
import type { RequestSpec } from './types.js';

export function renderRequest(spec: RequestSpec): string {
  const lines: string[] = [];

  lines.push(`${spec.method} ${spec.url}`);

  const headers = redactHeaders(spec.headers ?? {});
  for (const [key, value] of Object.entries(headers)) {
    lines.push(`${key}: ${value}`);
  }

  if (spec.body !== undefined) {
    lines.push('');
    lines.push(JSON.stringify(redactObject(spec.body), null, 2));
  }

  return lines.join('\n');
}
