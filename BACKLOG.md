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
   - Status: partially completed. High-value parsing paths now emit structured warnings, and conversion APIs can return captured diagnostics. Remaining work is deeper coverage across the remaining silent/console-only paths.

## P1

3. Chart fidelity improvements
   - Goal: move from editable placeholder charts toward better source reconstruction.
   - Deliverables: chart titles, legend presence, axis labels, more series metadata, and broader BIFF record handling.
   - Exit criteria: common embedded Office charts preserve more than category/value grids.

4. Theme interpretation beyond raw extraction
   - Goal: use extracted theme XML to influence generated formatting instead of only preserving the payload.
   - Deliverables: parsed color scheme, font scheme, and theme-aware color resolution.
   - Exit criteria: theme-backed formatting in converted DOCX matches source documents more closely.

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
