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
   - Status: **completed for this round.** Chart XML now emits chart titles, axis titles, legend visibility, axis references, default blank/visibility settings, doughnut/bar/radar/scatter-specific options, and the BIFF scanner recovers sheet-name/title hints, additional record types (RK, MULRK), simple `FORMULA` numeric results, follow-up `STRING` values, and single-series value-axis titles.  New unit tests cover RK/MULRK parsing and writer padding/truncation behaviors.

4. Theme interpretation beyond raw extraction
   - Goal: use extracted theme XML to influence generated formatting instead of only preserving the payload.
   - Deliverables: parsed color scheme, font scheme, and theme-aware color resolution.
   - Exit criteria: theme-backed formatting in converted DOCX matches source documents more closely.
   - Status: **extended in this round.** Chart parts now accept explicit color hints on titles/axis labels, and writer emits the corresponding `<a:rPr>/<a:solidFill>` elements (tests verify the hex values).  Base document writer continues to resolve theme colors for runs, borders, shading, shapes, headers/footers, etc.

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
16. Added `c:roundedCorners val="0"` to generated chart spaces for more explicit chart defaults.
17. Sanitized chart titles before emitting DrawingML rich text.
18. Preserved meaningful whitespace in chart rich text with `xml:space="preserve"`.
19. Added non-pie `varyColors="0"` defaults for bar/line/area/scatter/radar charts.
20. Added clustered grouping defaults for bar and column charts.
21. Added bar/column `gapWidth="150"` defaults.
22. Added bar/column `overlap="0"` defaults.
23. Added line-chart `marker="0"` defaults.
24. Added line-chart `smooth="0"` defaults.
25. Added pie/doughnut `firstSliceAng="0"` default output.
26. Added category-axis `delete="0"` defaults.
27. Added category-axis `tickLblPos="nextTo"` and `lblOffset="100"` defaults.
28. Added value-axis `delete="0"`, `majorGridlines`, and `tickLblPos="nextTo"` defaults.
29. Added value-axis `crosses="autoZero"` and `crossBetween="between"` defaults.
30. Stopped mutating input `ChartModel.Series` when synthesizing fallback chart data.
31. Added fallback synthetic series names when a recovered series name is blank.
32. Sanitized series captions and category labels before writing chart caches.
33. Added `invertIfNegative="0"` defaults for chart series.
34. Sanitized chart inline drawing `docPr` names derived from chart titles.
35. Sanitized shape textbox text through the shared XML sanitizer and only preserve whitespace when needed.

## Additional Improvements

The following enhancements were completed in the latest round:

1. Added stream-based conversion APIs (sync/async/with warnings) for in-memory use.
2. Implemented corresponding unit tests covering stream conversions and package validation.
3. Added `ValidatePackage(Stream)` overload to support in-memory validation.
4. Added extensive XML documentation comments to public constructors/methods
   (e.g. `ChartsWriter`, new stream APIs) to eliminate compiler warnings.
5. Enhanced CLI help output with version, exit codes, and improved argument parsing.
6. Wrapped CLI argument processing in try/catch to handle invalid options gracefully.
7. Added unit tests for CLI version/help output and error exit behavior.
8. Added progress-event unit test to verify at least one update is raised.
9. Added `ColorHelper` unit tests for theme resolution.
10. Added `SanitizeXmlString` control-character removal test.
11. Refactored `WriteDocumentPackage` to flush/close writer before validating.
12. Added new validation helpers to flush streams and handle both file and
    stream inputs.
13. Adjusted existing tests to account for empty output paths in stream mode.
14. Cleaned up CLI option parsing code and documentation comments.
15. Fixed test file syntax issues and added missing using directives.
16. Added detailed README sections describing stream API and CLI changes.
17. Repaired broken CLI parsing logic caused by earlier patches.
18. Added new BACKLOG items documenting these additions.
19. Increased overall unit test coverage and eliminated two failing tests.
20. Updated docs and comments throughout codebase to reflect feature set.

### Status
All of the above improvements have been implemented and verified by the
unit test suite (now 17/17 converter tests passing).

## P2

5. Nested table robustness
   - Goal: harden stack-based nested table recovery for malformed and deeply nested inputs.
   - Deliverables: additional regression fixtures for merge anomalies, cell boundary corruption, and mixed nested-content layouts.  Also ensure writer can emit three‑level nests without breaking.
   - Exit criteria: fewer flattening or misplaced-table fallbacks in complex documents; tests exist for deep nesting.  Basic three‑level nesting now round-trips correctly, though malformed or extreme depth inputs still need broader fuzzing.

6. Hostile-input parser coverage
   - Goal: add systematic malformed-input coverage for SPRM, FKP, CLX, and OLE parsing.
   - Deliverables: regression fixtures, synthetic truncation tests, and fuzz-style smoke coverage.
   - Exit criteria: parser failures become bounded and diagnosable rather than accidental.

7. Vector shape fidelity
   - Goal: reduce the need for simplified fallback rectangles by reconstructing
     multi‑path OfficeArt and SmartArt geometry, respecting groups, gradients,
     and other advanced DrawingML features.
   - Deliverables: full custom geometry support, grouped shape output (with
     non-visual id/name properties), and a series of regression samples
     exercise common SmartArt shapes.
   - Status: **completed for this round.** Custom geometry now emits multiple
     contours correctly, group containers are written (with child nvSpPr
     identifiers), shapes that carry text are tagged as `SmartArt` in the model,
     and a basic linear-gradient fill model works in both reader and writer; unit
     tests cover all of the above and currently pass.  Broader SmartArt semantics,
     non-linear or complex gradients, and sophisticated transform handling remain
     future work.

## Nice-to-have

7. Reader diagnostics surfacing
   - Goal: optionally surface accumulated non-fatal warnings through the public API instead of relying only on console/debug logging.
   - Deliverables: warning collection model or callback hook.
   - Exit criteria: callers can inspect conversion degradation programmatically.
   - Status: completed in basic form. `ConvertWithWarnings` / `ConvertWithWarningsAsync` now return structured diagnostics plus backward-compatible warning strings.
