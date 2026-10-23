import { analyzeBindings, type BoundExpression, closest, type ElementAnalysis } from '../analysis/bindings.js';
import { isTemplateLabel, TEMPLATE_LABEL_PREFIX } from '../model/projectData.js';
import type { DocentricTemplate } from '../model/types.js';

export type Severity = 'error' | 'warning' | 'info';

export interface Finding {
  code: string;
  severity: Severity;
  message: string;
  elementId?: string;
  elementType?: string;
  part?: string;
  /** What to do about it — written for the agent that will act on the finding. */
  hint?: string;
}

/**
 * Rule catalogue. Codes are stable: add new ones at the end, never renumber.
 * Severity reflects what the Docentric engine does at print time:
 *   error   → the element renders wrong or the print fails
 *   warning → probably a mistake, the document still prints
 *   info    → worth knowing, nothing is broken
 */
export const RULES: Record<string, { severity: Severity; title: string }> = {
  DCX001: { severity: 'error', title: 'Content control references an element that is not defined' },
  DCX002: { severity: 'warning', title: 'Element is defined but not placed in the document' },
  DCX003: { severity: 'error', title: 'ChildElements references an element that is not defined' },
  DCX004: { severity: 'warning', title: 'Content control title (alias) differs from the element type' },
  DCX005: { severity: 'error', title: 'Binding uses an unknown data source' },
  DCX006: { severity: 'error', title: "Source='.' outside any List/Group" },
  DCX007: { severity: 'error', title: 'Binding path not found in the data source schema' },
  DCX008: { severity: 'error', title: 'Template label is used but not defined' },
  DCX009: { severity: 'warning', title: 'Template label has no text for a template language' },
  DCX010: { severity: 'info', title: 'Template label is defined but never used' },
  DCX011: { severity: 'info', title: 'D365 label resolved at run time — not checkable offline' },
  DCX012: { severity: 'error', title: 'var-element-value() references an undeclared variable' },
  DCX013: { severity: 'warning', title: 'Element is placed more than once' },
  DCX014: { severity: 'warning', title: 'Content controls carry more than one project id' },
  DCX015: { severity: 'warning', title: 'Element has an empty binding' },
  DCX016: { severity: 'error', title: 'Binding expression cannot be parsed' },
  DCX017: { severity: 'warning', title: 'List/Group binding does not select an element collection' },
  DCX018: { severity: 'warning', title: 'Data source has no XSD — its bindings cannot be checked' },
  DCX019: { severity: 'warning', title: 'Data source schema is provisional (generated, not from a DDSP)' },
};

const NEEDS_BINDING = new Set(['Field', 'If', 'List', 'Group', 'Var', 'Image', 'SummaryField']);

export interface ValidationResult {
  findings: Finding[];
  counts: Record<Severity, number>;
  checkedPaths: number;
}

export function validateTemplate(template: DocentricTemplate): ValidationResult {
  const findings: Finding[] = [];
  const add = (code: string, f: Omit<Finding, 'code' | 'severity'>) =>
    findings.push({ code, severity: RULES[code]?.severity ?? 'warning', ...f });

  const { elements, contentControls } = template;

  // --- Placement: content controls ↔ element registry -------------------------
  const placements = new Map<string, number>();
  for (const cc of contentControls) {
    const id = cc.elementId as string;
    placements.set(id, (placements.get(id) ?? 0) + 1);
    const el = elements.get(id);
    if (!el) {
      add('DCX001', {
        elementId: id,
        part: cc.part,
        message: `Content control '${cc.tag}' (${cc.alias ?? 'no title'}) in ${cc.part} has no element ${id} in ProjectData.`,
        hint: 'Usually content copied from another template. Re-insert the element with the Docentric Designer, or delete the content control.',
      });
    } else if (cc.alias && cc.alias !== el.type) {
      add('DCX004', {
        elementId: id,
        elementType: el.type,
        part: cc.part,
        message: `Content control '${cc.tag}' is titled '${cc.alias}' but element ${id} is a ${el.type}.`,
      });
    }
  }
  for (const [id, n] of placements) {
    if (n > 1) {
      add('DCX013', {
        elementId: id,
        elementType: elements.get(id)?.type,
        message: `Element ${id} is placed ${n} times; editing one copy in the Designer changes all of them.`,
      });
    }
  }
  const projectIds = new Set(contentControls.map((c) => c.projectId));
  if (projectIds.size > 1) {
    add('DCX014', {
      message: `Content controls use ${projectIds.size} project ids (${[...projectIds].join(', ')}).`,
      hint: 'Content pasted from another Docentric template keeps its original project id. Check those elements.',
    });
  }

  const referenced = new Set<string>();
  for (const el of elements.values()) {
    for (const childId of el.childIds) {
      referenced.add(childId);
      if (!elements.has(childId)) {
        add('DCX003', {
          elementId: el.id,
          elementType: el.type,
          message: `${el.type} ${el.id} lists child ${childId}, which is not defined.`,
        });
      }
    }
  }
  for (const el of elements.values()) {
    if (!placements.has(el.id)) {
      add('DCX002', {
        elementId: el.id,
        elementType: el.type,
        message: `${el.type} ${el.id}${el.name ? ` '${el.name}'` : ''} is in ProjectData but no content control shows it.`,
        hint: 'Leftover from deleted content. Harmless, but it still appears in the Designer element tree.',
      });
    }
  }

  // --- Data sources ----------------------------------------------------------
  for (const ds of template.dataSources) {
    if (!ds.root) add('DCX018', { message: `Data source '${ds.name}' has no XSD.` });
  }
  for (const name of template.provisionalDataSources) {
    add('DCX019', {
      message: `Data source '${name}' has a provisional schema generated by the docentric skill, not one exported by D365.`,
      hint: 'Generate the DDSP in D365 and load it in the Docentric Designer (Data sources → replace/update) before deploying; then re-run validate.',
    });
  }

  // --- Bindings ----------------------------------------------------------------
  const analysis = analyzeBindings(template);
  const varNames = new Set(
    [...elements.values()].filter((e) => e.type === 'Var' && e.name).map((e) => e.name as string),
  );
  const labelUses = new Map<string, string[]>();
  const useLabel = (code: string, elementId: string) => {
    const list = labelUses.get(code) ?? [];
    list.push(elementId);
    labelUses.set(code, list);
  };
  let checkedPaths = 0;

  for (const a of analysis.values()) {
    const el = a.element;
    if (el.labelCode) useLabel(el.labelCode, el.id);
    if (NEEDS_BINDING.has(el.type) && (!el.binding || el.binding.path.trim() === '')) {
      add('DCX015', { elementId: el.id, elementType: el.type, message: `${el.type} ${el.id} has no binding path.` });
    }
    for (const [kind, bound] of [
      ['binding', a.value],
      ['aggregate binding', a.aggregate],
    ] as const) {
      if (!bound || bound.binding.path.trim() === '') continue;
      checkedPaths += checkBound(a, kind, bound, add);
      for (const code of bound.info.labelCodes) useLabel(code, el.id);
      for (const v of bound.info.varRefs) {
        if (!varNames.has(v)) {
          add('DCX012', {
            elementId: el.id,
            elementType: el.type,
            message: `${el.type} ${el.id} reads variable '${v}', but no Var element is named '${v}'.`,
            hint: varNames.size
              ? `Declared variables: ${[...varNames].join(', ')}`
              : 'The template declares no Var elements.',
          });
        }
      }
    }
    if (
      CONTEXT_PROVIDER(el.type) &&
      a.value?.context &&
      !a.value.selects &&
      a.value.paths.every((p) => p.resolution.status === 'ok')
    ) {
      add('DCX017', {
        elementId: el.id,
        elementType: el.type,
        message: `${el.type} ${el.id} binds '${el.binding?.path}', which is not a plain path to an element; children using Source='.' cannot be checked.`,
      });
    }
  }

  // --- Labels ------------------------------------------------------------------
  const own = template.labels.filter(isTemplateLabel);
  const defined = new Map(own.map((l) => [l.code, l]));
  for (const [code, users] of labelUses) {
    if (code.startsWith(TEMPLATE_LABEL_PREFIX)) {
      if (!defined.has(code)) {
        const near = closest(
          code.slice(TEMPLATE_LABEL_PREFIX.length),
          own.map((l) => l.code.slice(TEMPLATE_LABEL_PREFIX.length)),
        );
        add('DCX008', {
          elementId: users[0],
          message: `Label '${code}' is used by element(s) ${users.join(', ')} but not defined in the template labels.`,
          hint:
            (near.length ? `Did you mean ${near.map((n) => `'${TEMPLATE_LABEL_PREFIX}${n}'`).join(', ')}? ` : '') +
            'Otherwise add it in Designer → Labels (codes are case-sensitive).',
        });
      }
    } else if (code.startsWith('@')) {
      add('DCX011', {
        elementId: users[0],
        message: `Label '${code}' (element(s) ${users.join(', ')}) comes from D365.`,
        hint: 'Check it exists in every printing language with the d365fo-mcp server (`labels` tool).',
      });
    }
  }
  const languages = template.languages.length
    ? template.languages
    : template.defaultLanguage
      ? [template.defaultLanguage]
      : [];
  const unused: string[] = [];
  for (const label of own) {
    const missing = languages.filter((lang) => !(label.texts[lang] ?? '').trim());
    if (missing.length) {
      add('DCX009', { message: `Label '${label.code}' has no text for: ${missing.join(', ')}.` });
    }
    if (!labelUses.has(label.code)) unused.push(label.code);
  }
  if (unused.length) {
    // One finding, not one per label: templates copied from a base often carry hundreds.
    const shown = unused.slice(0, 15).join(', ');
    add('DCX010', {
      message: `${unused.length} of ${own.length} template labels are never used: ${shown}${unused.length > 15 ? ', …' : ''}.`,
      hint: 'Unused labels are harmless; remove them in Designer → Labels to keep translation work small.',
    });
  }

  const order: Record<Severity, number> = { error: 0, warning: 1, info: 2 };
  findings.sort((x, y) => order[x.severity] - order[y.severity] || x.code.localeCompare(y.code));
  const counts: Record<Severity, number> = { error: 0, warning: 0, info: 0 };
  for (const f of findings) counts[f.severity]++;
  return { findings, counts, checkedPaths };
}

const CONTEXT_PROVIDER = (type: string) => type === 'List' || type === 'Group';

function checkBound(
  a: ElementAnalysis,
  kind: string,
  bound: BoundExpression,
  add: (code: string, f: Omit<Finding, 'code' | 'severity'>) => void,
): number {
  const el = a.element;
  const where = `${el.type} ${el.id} ${kind}`;
  if (bound.info.error) {
    add('DCX016', {
      elementId: el.id,
      elementType: el.type,
      message: `${where} '${bound.binding.path}': ${bound.info.error}.`,
    });
    return 0;
  }
  if (bound.contextError) {
    const code =
      bound.contextError.code === 'unknown-source'
        ? 'DCX005'
        : bound.contextError.code === 'no-context'
          ? 'DCX006'
          : null;
    if (code)
      add(code, { elementId: el.id, elementType: el.type, message: `${where}: ${bound.contextError.message}.` });
    return 0;
  }
  let checked = 0;
  for (const p of bound.paths) {
    if (p.resolution.status === 'skipped') continue;
    checked++;
    if (p.resolution.status === 'missing') {
      const r = p.resolution;
      const ctx = bound.context;
      const base = ctx ? ctx.chain.map((n) => n.name).join('/') : '?';
      const at = r.resolvedText ? `${base}/${r.resolvedText}` : base;
      add('DCX007', {
        elementId: el.id,
        elementType: el.type,
        message: `${where} '${bound.binding.path}': '${r.step}' does not exist under ${ctx?.dataSource.name}:${at}.`,
        hint: r.suggestions.length
          ? `Did you mean ${r.suggestions.map((s) => `'${s}'`).join(', ')}?`
          : 'If the field was just added in X++, regenerate the DDSP in D365 and update the template data source.',
      });
    }
  }
  return checked;
}

export function renderFindings(result: ValidationResult, minSeverity: Severity = 'info'): string {
  const rank: Record<Severity, number> = { error: 0, warning: 1, info: 2 };
  const shown = result.findings.filter((f) => rank[f.severity] <= rank[minSeverity]);
  const lines = [
    `Validation: ${result.counts.error} error(s), ${result.counts.warning} warning(s), ${result.counts.info} info — ${result.checkedPaths} binding path(s) checked against the schema.`,
  ];
  if (shown.length < result.findings.length)
    lines.push(`(showing severity ≥ ${minSeverity}: ${shown.length} of ${result.findings.length})`);
  for (const f of shown) {
    lines.push(`- [${f.severity}] ${f.code} ${f.message}`);
    if (f.hint) lines.push(`    ↳ ${f.hint}`);
  }
  return lines.join('\n');
}
