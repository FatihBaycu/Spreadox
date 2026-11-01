# Changelog

All notable changes to the **Spreadox** library will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.0] - 2026-11-01

### Changed
- **Official Package Rebrand (`Spreadox`):** Launched standalone NuGet package ID `Spreadox` (`dotnet add package Spreadox`) for optimal discoverability and clean naming across search engines, AI assistants, and developer tooling.
- **AI & SEO Metadata Optimization:** Enhanced NuGet package description and keywords targeting high-intent spreadsheet, reporting, and OOXML workflows (`excel`, `xlsx`, `spreadsheet`, `openxml`, `pdf`, `chart`, `streaming`, `vector`).

---

## [0.9.2] - 2026-11-01

### Added
- **Advanced Resilience & Robustness Suite (185 Tests Total, +29 Tests):**
  - **Excel Differential Testing:** Verified consistent semantics across Spreadox DOM, SAX Streaming, and Microsoft OpenXML SDK DOM engines with round-trip mutation safety (`ExcelDifferentialTestingTests`).
  - **Crash & Interruption Recovery:** Verified that incomplete or prematurely aborted stream exports are strictly rejected upon reopen, and unmanaged file locks are immediately released (`CrashRecoveryTests`).
  - **Memory Pressure & GC Stability:** Validated sustained exports under 128 MB artificial memory pressure (`GC.AddMemoryPressure`) with bounded heap growth (`MemoryPressureTests`).
  - **Concurrent Stress Testing:** 100 concurrent worker tasks generating 200 documents simultaneously in ~2.8s (avg 28 ms, p95 = 52 ms) with 100% data isolation and 0 OS handle leaks (`ConcurrentStressTests`).
  - **Formula Injection (DDE) Hardening:** 14 attack payloads (`=cmd`, `+`, `-`, `@`, leading whitespace/tabs) confirmed stored strictly as pure text, preventing spreadsheet code execution while preserving functional formulas (`FormulaInjectionRegressionTests`).
  - **ZIP & XML Security:** Validated protection against ZipSlip directory traversal, excessive ZIP entries, XXE injection (DTD disabled), and malformed XML (`ZipAndXmlSecurityTests`).
  - **PDF Font Cache Concurrency:** Validated thread-safe initialization and byte immutability across 100 concurrent PDF generations, eliminating ~250 MB of LOH memory bloat (`PdfFontCacheConcurrencyTests`).
  - **Reproducible Build Verification:** Deterministic build flags (`<Deterministic>true</Deterministic>`), SourceLink integration, and snupkg symbols verified (`ReproducibleBuildTests`).
- **Audit & Benchmark Documentation:**
  - Updated [docs/TEST_RESULTS.md](docs/TEST_RESULTS.md) with complete 185-test execution, memory profile, and benchmark matrix.

### Changed
- **Repository Rename:** Updated all repository, release notes, and documentation links to `https://github.com/FatihBaycu/Spreadox`.
- **MVC Sample Decoupling:** `ReportBuilderMvc` now consumes `Conq.Noble.Spreadox` directly as a NuGet package (`PackageReference`) rather than an in-solution project reference.
- **Docker Build Optimization:** Simplified `Dockerfile` layer caching for MVC app by removing redundant source project copy.

### Fixed
- **Invalid Package Ingestion Hardening (DEF-007):** `OpenXmlWorkbookReader` now validates package structure and throws `InvalidDataException` when reading non-Excel or empty ZIP packages without a valid `WorkbookPart` or `Sheets`.

---

## [0.9.1] - 2026-10-31

### Added
- **Comprehensive Production Readiness & Resilience Test Suite (156 Tests Total):**
  - **High-Volume Streaming Benchmarks:** Validated 10,000 to 1,000,000 rows x 10 cols with exact allocation and GC collection metrics.
  - **Concurrency & Stress Tests:** Verified data isolation across 1, 5, 10, 25, 50, and 100 parallel worker threads.
  - **Round-Trip Fidelity:** Spreadox $\rightarrow$ Microsoft OpenXML SDK modify $\rightarrow$ Spreadox reload; preserves formulas, styles, charts, images, and freeze panes.
  - **Property-Based Invariants:** 10,000 randomized coordinates and 2,000 2D ranges validating bidirectional formatting invariants (`Format` $\leftrightarrow$ `Parse`).
  - **Fuzz Testing:** Robust handling of corrupted byte streams, truncated ZIP packages, 50,000-character cell strings, and malformed font streams.
  - **Formula Engine & Security:** Validated SUM, AVERAGE, COUNT, IF, nested expressions, and CSV/Formula injection payload neutralization.
  - **Style Deduplication:** Automatic style deduplication collapsing 10,000 styled cells into minimal style records.
  - **Streaming Fault Tolerance:** `FaultyStream` disk-full recovery, stream disposal guarantees, and row monotonicity enforcement.
  - **Financial Decimal Fidelity:** Strict decimal precision for currency and financial calculations without floating-point drift.
  - **Excel Boundary Checks:** Enforced limits on rows (1,048,576), columns (16,384), cell text (32,767), and sheet names (31 chars).
  - **PDF Visual Geometry:** Exact A4 landscape geometry, UCL/CL/LCL control lines, and PDF text extractability via PdfPig.
  - **Long-Running Workloads:** 500 continuous mixed Excel/PDF production cycles with zero OS handle leaks.
  - **Culture Invariance:** Invariant (`.`) number serialization across Turkish, German, French, and Arabic locales.
  - **Public API Isolation:** Reflection verification confirming zero vendor type leaks (`DocumentFormat.OpenXml`, `PdfSharpCore`) in public APIs.
- **Workflow & Audit Documentation:**
  - Added [docs/TEST_RESULTS.md](docs/TEST_RESULTS.md) with complete 156-test execution, memory profile, and benchmark matrix.
  - Added release, versioning, and commit workflow rules to `.gemini/rules/` and `.cursor/rules/`.

### Fixed
- **XML 1.0 Control Character Sanitization (DEF-005):** Implemented `SanitizeXmlString` in `OpenXmlWorkbookWriter` and `StreamingPackageWriter` to strip illegal control characters (`\0`..`\x1F` except `\t`, `\r`, `\n`), preventing serialization crashes.
- **Coordinate Boundary Validation (DEF-001):** Enforced maximum 1,048,576 rows and 16,384 columns in `CellAddress.Parse`.
- **Worksheet Name Hardening (DEF-002):** Enforced Excel forbidden character checks (`\ / ? * [ ] :`), leading/trailing apostrophes (`'`), and duplicate sheet names in `StreamingPackageWriter` and `ExcelWorksheet`.
- **Cell Nullability & Cross-Conversion (DEF-003):** Fixed nullable warning CS8602 in `ExcelCell` and enabled cross-conversion between integer, double, and decimal representations.

### Performance
- **Streaming Row Writer Allocation Reduction (DEF-004):** Reused `ExcelStreamingRowWriter` instance per worksheet, eliminating 1,000,000 Gen0 allocations during 1M-row exports.
- **TrueType Font LOH Caching (DEF-006):** Added `ConcurrentDictionary<string, byte[]>` cache in `PdfFontResolver`, eliminating ~2.5 MB Large Object Heap (LOH) allocations per PDF generation.

---

## [0.9.0] - 2026-10-31

### Added
- **`PdfHeader` Support:** Added full running header customization on `PdfReport` with `Title`, `Text` (subtitle), `Logo` (`PdfImage`), `FontFamily`, `FontSize`, `Bold`, and `TitleColor` (`ExcelColor`).
- **Runtime TrueType (`.ttf`) Font Registration (`PdfFontRegistry`):** Added public API to dynamically register custom TrueType fonts at runtime via `PdfFontRegistry.Register(familyName, ttfBytes)` or file path.
- **Font Support in `PdfFooter`:** Added `FontFamily` property to `PdfFooter` to allow using registered custom fonts in running footers.
- **Report Studio MVC Sample (`samples/ReportBuilderMvc`):**
  - Interactive ASP.NET Core MVC web application with live chart builder.
  - Dynamic random data generator with configurable series count and inline data editing.
  - Per-series chart type picker (`ColumnClustered`, `BarClustered`, `Line`, `LineMarker`, `Area`, `Pie`, `Doughnut`, `Scatter`, `Radar`, `Combo`).
  - Runtime `.ttf` font file upload & live registration.
  - Single-page vs. separate-page layout options for both Excel (`.xlsx`) and PDF (`.pdf`) exports.
  - Scalable data-driven multi-language support (English default, Turkish) with a floating glassmorphic language switcher.
- **Visual Documentation:** Added high-resolution (2x Retina) vector PDF and Excel output screenshots under `docs/images/`.

### Changed
- **`PdfReportWriter` Spacing:** Increased `HeaderHeight` to 36pt, refined dynamic subtitle vertical offset, and added an 8pt (~3mm) breathing margin below the header divider line so it never overlaps subtitle text.
- **Documentation:** Updated `README.md` with visual preview galleries and custom font code walkthroughs.

---

## [0.8.0] - 2026-10-15

### Added
- **Excel DOM Workbook (`ExcelWorkbook`):**
  - Microsoft Open XML SDK based spreadsheet engine free of OpenXML public types.
  - Native OOXML DrawingML chart support (Column, Bar, Line, Scatter, Area, Pie, Doughnut, Radar, Bubble, Combo, BoxAndWhisker).
  - Cell formatting: Font, Fill, Border styles, Alignment, Number Formats.
  - Conditional formatting (Comparison rules and 3-color scales).
  - Merged cells, Freeze panes, and Auto-filter support.
- **High-Volume Streaming Engine (`ExcelStreamingWorkbook`):** Low-memory forward-only writer capable of flushing 100,000+ rows directly to disk.
- **Vector PDF Engine (`PdfReport`):**
  - A4 Landscape publication-ready PDF chart reports.
  - Vector-drawn charts from in-memory arrays (not OOXML charts).
  - Cover page support with metadata facts, company title, and date.
  - Running footer with page numbering format (`Page {0} / {1}`) and logo.
  - Right-hand metric summary sidebar (up to 28 rows).
  - Embedded SIL Open Font License Liberation Sans fonts.

