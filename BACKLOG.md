# Development Backlog

This backlog turns the current audit findings into a practical implementation order.

## P0

1. RC4 encrypted DOC end-to-end validation
   - Goal: remove ambiguity around WordDocument decryption boundaries.
   - Deliverables: real encrypted regression samples, pass/fail password cases, coverage for WordDocument/Table/Data/story streams, and assertions for text/image recovery.
   - Exit criteria: encrypted sample conversions stop depending on undocumented offset assumptions.

2. Silent parse failure diagnostics
   - Goal: replace silent best-effort catches in high-value parsing paths with structured warnings.
   - Deliverables: consistent logging for OLE extraction, OfficeArt/image scanning, chart scanning, and optional stream recovery.
   - Exit criteria: partial conversion failures are visible in logs without aborting the full conversion.
   - Status: partially completed. High-value parsing paths now emit structured warnings, conversion APIs can return captured diagnostics, the remaining direct reader-side `Console.WriteLine` warnings have been unified behind `Logger`, and several previously silent text/shading/font fallback branches now emit debug or warning traces. Remaining work is deeper classification across the remaining best-effort fallback branches.

## P1

3. Chart fidelity improvements
   - Goal: move from editable placeholder charts toward better source reconstruction.
   - Deliverables: chart titles, legend presence, axis labels, more series metadata, and broader BIFF record handling.
   - Exit criteria: common embedded Office charts preserve more than category/value grids.
   - Status: partially completed. Chart XML now emits chart titles, axis titles, legend visibility, axis references, default blank/visibility settings, doughnut/bar/radar/scatter-specific options, and the BIFF scanner recovers sheet-name/title hints, additional record types, simple `FORMULA` numeric results, follow-up `STRING` values, and single-series value-axis titles.

4. Theme interpretation beyond raw extraction
   - Goal: use extracted theme XML to influence generated formatting instead of only preserving the payload.
   - Deliverables: parsed color scheme, font scheme, and theme-aware color resolution.
   - Exit criteria: theme-backed formatting in converted DOCX matches source documents more closely.
   - Status: partially completed. Theme XML is now parsed into color/font metadata, default DOCX fonts prefer the extracted body theme fonts, and theme-referenced colors are emitted across body runs, hyperlinks, borders, shading, shapes, comments, footnotes, endnotes, and paragraph-based header/footer content with concrete RGB fallbacks where possible.

11. Propagated theme-aware run formatting through hyperlink runs in the main document writer.
12. Replaced remaining direct `ColorToHex` run-color fallbacks with theme-aware resolved color output.
13. Added single-series value-axis title inference from recovered series labels.
14. Turned several text/font/shading fallback silent catches into debug traces.
15. Elevated garbled footnote/endnote table-stream fallback to an explicit warning.

## This Round

1. Switched comment run output to the shared `RunPropertiesHelper` pipeline.
2. Added XML sanitization for comment text runs.
3. Switched footnote and endnote run output to the shared `RunPropertiesHelper` pipeline.
4. Extended theme-aware formatting to comments, footnotes, and endnotes.
5. Added BIFF `FORMULA` numeric result recovery.
6. Added BIFF follow-up `STRING` recovery for formula-backed labels.
7. Added doughnut/donut source-name chart type detection.
8. Added radar source-name chart type detection.
9. Replaced the remaining direct reader-side console warnings with structured logger warnings.
10. Downgraded noisy package/table trace logs from `Info` to `Debug` so normal library calls stay quiet.
11. Bound header/footer paragraph writing to document theme/style context instead of treating it as a context-free fragment.
12. Disabled external hyperlink emission in header/footer fragments where no dedicated part relationships are generated, falling back to plain themed text.
13. Reused shared XML text sanitization for simple header/footer text fallback.
14. Added chart `axId` references plus `autoTitleDeleted`/`plotVisOnly`/`dispBlanksAs` defaults for more stable Word chart parts.
15. Replaced the remaining `TableReader` debug-log cleanup bare catch with a debug trace.

## P2

5. Nested table robustness
   - Goal: harden stack-based nested table recovery for malformed and deeply nested inputs.
   - Deliverables: additional regression fixtures for merge anomalies, cell boundary corruption, and mixed nested-content layouts.
   - Exit criteria: fewer flattening or misplaced-table fallbacks in complex documents.

6. Hostile-input parser coverage
   - Goal: add systematic malformed-input coverage for SPRM, FKP, CLX, and OLE parsing.
   - Deliverables: regression fixtures, synthetic truncation tests, and fuzz-style smoke coverage.
   - Exit criteria: parser failures become bounded and diagnosable rather than accidental.

## Nice-to-have

7. Reader diagnostics surfacing
   - Goal: optionally surface accumulated non-fatal warnings through the public API instead of relying only on console/debug logging.
   - Deliverables: warning collection model or callback hook.
   - Exit criteria: callers can inspect conversion degradation programmatically.
   - Status: completed in basic form. `ConvertWithWarnings` / `ConvertWithWarningsAsync` now return structured diagnostics plus backward-compatible warning strings.
