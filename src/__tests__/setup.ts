import { configure } from '@testing-library/dom'
import '@testing-library/jest-dom/vitest'

/**
 * Raise Testing Library's async-utility budget from its 1000 ms default.
 *
 * Every `waitFor`/`findBy*` in the suite shares this one timeout, and 417 of
 * the suite's 420 `waitFor` calls take it as-is (only three pass an explicit
 * `timeout`). 1000 ms is generous when a file runs alone and marginal when 281
 * test files run in parallel: `App.dirtyState.characterization.test.tsx`'s wait
 * for `.markdown-body` — the first-ever mount of the `React.lazy()`
 * `MarkdownRenderer` chunk (~500 KB, PERF-01) — passes in isolation in under
 * 4 s total but intermittently times out inside a full `npm test` run, and CI's
 * machine has less headroom than a dev box, not more.
 *
 * This is a ceiling on how long a wait MAY take before it's called a failure,
 * not a delay anything pays: a satisfied condition resolves on the next poll
 * either way, so raising it costs passing runs nothing and only buys a loaded
 * or slow machine room to finish. Deliberately global rather than a
 * per-`waitFor` timeout at the one site that flaked, since the cause (lazy
 * chunk + parallel load) is shared by every lazily-mounted viewer's tests.
 */
configure({ asyncUtilTimeout: 5000 })
