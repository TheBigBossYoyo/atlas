/**
 * D23-PERF — document-wide metadata `PageView` needs per paragraph
 * (tracked-change/hyperlink run metadata, bookmark anchor names), split out
 * of `PageView.tsx` into its own module so `PageStack.tsx` can compute it
 * ONCE per document and hand the result to every page (see `PageStack.tsx`'s
 * own doc comment) instead of every `PageView` re-walking the whole document
 * itself. A plain function/constant module also satisfies the
 * `react-refresh/only-export-components` lint rule, which only allows a
 * component file to export components.
 */
import type {
  Document,
  Hyperlink,
  HyperlinkChild,
  ParagraphChild,
} from '../model/document';
import type { Relationship } from '../parser/relationships';
import type { RevisionRenderKind } from './style';

/** Prefix for a bookmark's rendered anchor `id` — an internal hyperlink's
 * `href` points at `#${BOOKMARK_ANCHOR_ID_PREFIX}${anchor}`, letting the
 * browser's own same-page fragment navigation do the scrolling (D6). */
export const BOOKMARK_ANCHOR_ID_PREFIX = 'docx-bookmark-';

type RevisionRunMeta = {
  readonly kind: RevisionRenderKind;
  readonly author?: string;
};

type HyperlinkRunMeta = {
  readonly href: string;
  readonly isExternal: boolean;
  readonly tooltip?: string;
};

export type RunMeta = {
  readonly revision?: RevisionRunMeta;
  readonly hyperlink?: HyperlinkRunMeta;
};

/**
 * Walks every paragraph once, producing per-run-index metadata (tracked-
 * change revision AND/OR enclosing hyperlink target) in the exact same
 * flattened run order paginate.ts's `collectParagraphRuns`/
 * `collectHyperlinkRuns` produce `LineItem.runIndex` values in — run, then
 * a hyperlink's own runs, then a revision's own runs, all in document
 * order. The two concerns share one walk (rather than two independent
 * ones) so they can never drift apart on what "the i-th run" means.
 */
export function collectRunMetaByParagraph(
  document: Document,
  relationships: ReadonlyArray<Relationship>,
): ReadonlyMap<string, ReadonlyArray<RunMeta | undefined>> {
  const paragraphs = new Map<string, ReadonlyArray<RunMeta | undefined>>();

  for (const section of document.sections) {
    section.blocks.forEach((block, blockIndex) => {
      if (block.kind !== 'paragraph') {
        return;
      }

      const runs: Array<RunMeta | undefined> = [];
      appendParagraphChildRuns(block.children, runs, relationships, {});
      paragraphs.set(String(blockIndex), runs);
    });
  }

  return paragraphs;
}

function appendParagraphChildRuns(
  children: ReadonlyArray<ParagraphChild>,
  runs: Array<RunMeta | undefined>,
  relationships: ReadonlyArray<Relationship>,
  context: RunMeta,
): void {
  for (const child of children) {
    if (child.kind === 'run') {
      runs.push(hasRunMeta(context) ? context : undefined);
      continue;
    }

    if (child.kind === 'hyperlink') {
      const hyperlink = resolveHyperlinkMeta(child, relationships);
      appendHyperlinkChildRuns(child.children, runs, {
        ...context,
        ...(hyperlink !== undefined ? { hyperlink } : {}),
      });
      continue;
    }

    if (child.kind === 'ins-revision' || child.kind === 'del-revision') {
      appendParagraphChildRuns(getRevisionChildren(child), runs, relationships, {
        ...context,
        revision: {
          kind: child.kind === 'ins-revision' ? 'ins' : 'del',
          ...(child.author !== undefined ? { author: child.author } : {}),
        },
      });
    }
  }
}

function appendHyperlinkChildRuns(
  children: ReadonlyArray<HyperlinkChild>,
  runs: Array<RunMeta | undefined>,
  context: RunMeta,
): void {
  for (const child of children) {
    if (child.kind === 'run') {
      runs.push(hasRunMeta(context) ? context : undefined);
    }
  }
}

function hasRunMeta(context: RunMeta): boolean {
  return context.revision !== undefined || context.hyperlink !== undefined;
}

function getRevisionChildren(child: Extract<ParagraphChild, { kind: 'ins-revision' | 'del-revision' }>): ReadonlyArray<ParagraphChild> {
  return child.children as ReadonlyArray<ParagraphChild>;
}

/**
 * External hyperlink targets come straight from the (untrusted) document's
 * own relationship parts — a crafted `.docx` could set one to
 * `javascript:`/`vbscript:`/`data:` etc. Chromium already refuses to
 * execute a `javascript:` URI opened via `target="_blank"`, and Electron's
 * own `setWindowOpenHandler` additionally allow-lists http/https before
 * calling `shell.openExternal` (see `electron/main.cjs`'s
 * `isAllowedExternalScheme`) — but this renders the raw string as a
 * literal `href` regardless, so it's still worth validating at the source
 * rather than depending solely on those other layers. Mirrors this
 * codebase's existing convention (`OdtViewer.tsx`'s DOMPurify sanitization,
 * `main.cjs`'s own scheme allow-list) of never trusting a URL scheme from
 * file content.
 */
const SAFE_HYPERLINK_SCHEMES: ReadonlySet<string> = new Set(['http:', 'https:', 'mailto:']);

function isSafeHyperlinkHref(href: string): boolean {
  try {
    // A base is required for a protocol-relative or bare host string; a
    // genuinely relative (schemeless) target is not a scheme we allow, so
    // this intentionally has no fallback base that would let one through.
    return SAFE_HYPERLINK_SCHEMES.has(new URL(href).protocol);
  } catch {
    return false;
  }
}

/**
 * Resolves a `Hyperlink` node to a renderable target: an external URL via
 * the document's relationships (only when the relationship is genuinely
 * marked `TargetMode="External"`, matching how DOCX generators — including
 * the one that built this wave's corpus fixture — author external
 * hyperlinks), or an internal same-document anchor via `w:anchor`
 * (resolved against a bookmark's rendered `id`, see
 * `BOOKMARK_ANCHOR_ID_PREFIX`). Returns `undefined` for a hyperlink this
 * viewer can't safely resolve (e.g. a relationship id with no matching
 * External relationship, or an unsafe URL scheme — see
 * `isSafeHyperlinkHref`).
 */
function resolveHyperlinkMeta(
  hyperlink: Hyperlink,
  relationships: ReadonlyArray<Relationship>,
): HyperlinkRunMeta | undefined {
  if (hyperlink.relationshipId !== undefined) {
    const relationship = relationships.find((candidate) => candidate.id === hyperlink.relationshipId);
    if (
      relationship === undefined ||
      relationship.targetMode !== 'External' ||
      !isSafeHyperlinkHref(relationship.target)
    ) {
      return undefined;
    }
    return {
      href: relationship.target,
      isExternal: true,
      ...(hyperlink.tooltip !== undefined ? { tooltip: hyperlink.tooltip } : {}),
    };
  }

  if (hyperlink.anchor !== undefined) {
    return {
      href: `#${BOOKMARK_ANCHOR_ID_PREFIX}${hyperlink.anchor}`,
      isExternal: false,
      ...(hyperlink.tooltip !== undefined ? { tooltip: hyperlink.tooltip } : {}),
    };
  }

  return undefined;
}

export function collectBookmarkNamesByParagraph(document: Document): ReadonlyMap<string, ReadonlyArray<string>> {
  const paragraphs = new Map<string, ReadonlyArray<string>>();

  for (const section of document.sections) {
    section.blocks.forEach((block, blockIndex) => {
      if (block.kind !== 'paragraph') {
        return;
      }

      const names = collectParagraphBookmarkNames(block.children);
      if (names.length > 0) {
        paragraphs.set(String(blockIndex), names);
      }
    });
  }

  return paragraphs;
}

function collectParagraphBookmarkNames(children: ReadonlyArray<ParagraphChild>): ReadonlyArray<string> {
  const names: string[] = [];

  for (const child of children) {
    if (child.kind === 'bookmark' && child.boundary === 'start' && child.name !== undefined) {
      names.push(child.name);
      continue;
    }

    if (child.kind === 'hyperlink') {
      for (const hyperlinkChild of child.children) {
        if (
          hyperlinkChild.kind === 'bookmark' &&
          hyperlinkChild.boundary === 'start' &&
          hyperlinkChild.name !== undefined
        ) {
          names.push(hyperlinkChild.name);
        }
      }
    }
  }

  return names;
}
