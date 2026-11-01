# Spreadox

[![NuGet](https://img.shields.io/nuget/v/Spreadox.svg)](https://www.nuget.org/packages/Spreadox)
[![CI](https://github.com/FatihBaycu/Spreadox/actions/workflows/ci.yml/badge.svg)](https://github.com/FatihBaycu/Spreadox/actions/workflows/ci.yml)
[![Live Demo](https://img.shields.io/badge/Live%20Demo-Vercel-black?logo=vercel)](https://spreadox-report-mvc.vercel.app/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/FatihBaycu/Spreadox/blob/main/LICENSE)

High-level, dependency-isolated C# library for **native Excel (.xlsx) workbooks** and **vector A4 landscape PDF chart reports**.

- **Excel:** Built on Microsoft Open XML SDK (DocumentFormat.OpenXml). Open XML types remain internal so your code stays clean. Supports cell styles, formulas, borders, colors, pictures, and native OOXML DrawingML charts.
- **PDF:** Sharp, scalable vector charts drawn directly from in-memory arrays (powered by PdfSharpCore). Includes cover pages, customizable running headers/footers with logos, metric summaries, and runtime TrueType (.ttf) font registration.

| | |
|---|---|
| **Package** | [Spreadox](https://www.nuget.org/packages/Spreadox) |
| **Namespace** | Conq.Noble.Spreadox |
| **Target** | netstandard2.0 (.NET Framework 4.8+, .NET 6/8/9+) |
| **Live Web Demo** | [https://spreadox-report-mvc.vercel.app/](https://spreadox-report-mvc.vercel.app/) |
| **QA & Benchmarks** | [docs/TEST_RESULTS.md](docs/TEST_RESULTS.md) (185 tests, 1M rows in 22s) |
| **Release Notes** | [CHANGELOG.md](https://github.com/FatihBaycu/Spreadox/blob/main/CHANGELOG.md) |
| **License** | [MIT](https://github.com/FatihBaycu/Spreadox/blob/main/LICENSE) |

---

## Visual Output Examples

### Interactive Web Studio (ReportBuilderMvc)

Configure series data, chart types, corporate cover page, headers/footers, and export directly to PDF or Excel.

[![Spreadox Report Studio](https://raw.githubusercontent.com/FatihBaycu/Spreadox/refs/heads/main/docs/images/spreadox-report-studio.png)](https://spreadox-report-mvc.vercel.app/)

> 🌐 **Try it live on Vercel:** [https://spreadox-report-mvc.vercel.app/](https://spreadox-report-mvc.vercel.app/)

### Vector PDF Reports (PdfReport & MVC Studio)

Vector charts rendered crisply at any zoom level, with running headers, logos, summary sidebars, and control lines.

| MVC Studio Cover Page | Radar Chart & Header (from MVC Studio) |
|:---:|:---:|
| ![PDF Cover](https://raw.githubusercontent.com/FatihBaycu/Spreadox/refs/heads/main/docs/images/mvc-pdf-cover.png) | ![PDF Radar Chart](https://raw.githubusercontent.com/FatihBaycu/Spreadox/refs/heads/main/docs/images/mvc-pdf-chart-radar.png) |

| Station Measurements (Line + UCL/LCL) | Pareto Combo (Bar + Secondary Line) |
|:---:|:---:|
| ![PDF Line Chart](https://raw.githubusercontent.com/FatihBaycu/Spreadox/refs/heads/main/docs/images/pdf-sample-chart-line.png) | ![PDF Pareto Chart](https://raw.githubusercontent.com/FatihBaycu/Spreadox/refs/heads/main/docs/images/pdf-sample-chart-pareto.png) |

| Bounce Rate (Pie Chart & Summary) | Bar Chart (with Summary Metrics) |
|:---:|:---:|
| ![PDF Pie Chart](https://raw.githubusercontent.com/FatihBaycu/Spreadox/refs/heads/main/docs/images/mvc-pdf-chart-pie.png) | ![PDF Bar Chart](https://raw.githubusercontent.com/FatihBaycu/Spreadox/refs/heads/main/docs/images/mvc-pdf-chart-bar.png) |

### Native Excel Spreadsheets (ExcelWorkbook)

Standard OOXML DrawingML charts linked to spreadsheet cell ranges with native formatting and styles.

| Excel Spreadsheet with Combo Chart and Summary Grid |
|:---:|
| ![Excel Report Sample](https://raw.githubusercontent.com/FatihBaycu/Spreadox/refs/heads/main/docs/images/excel-sample-report.png) |

---

## Install via NuGet

```bash
dotnet add package Spreadox
```

Package Manager:
```powershell
Install-Package Spreadox
```

---

## Quick Usage: Excel Workbook

```csharp
using Conq.Noble.Spreadox;

using var wb = ExcelWorkbook.Create();
var sheet = wb.Worksheets[0];
sheet.Name = "Demo";

// Formatted cells
sheet["A1"].SetValue("hello");
sheet["A1"].Style.Font.Bold = true;
sheet["B1"].SetValue(0.25);
sheet["B1"].Style.NumberFormat = NumberFormats.Percent;
sheet.SetFormula("D1", "B1*100", cachedValue: 25);
sheet.FreezePanes(1);

// Add clustered column chart
var chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, "Demo chart");
chart.AddSeries("B2:B10", "A2:A10", name: "Values");
chart.SetPosition(2, 4, 16, 12);

// Add combo chart (Columns on primary axis, Line on secondary axis)
var combo = sheet.Charts.Add(ExcelChartType.Combo, "Combo");
combo.AddSeries("B2:B10", "A2:A10", name: "Bars").AsChartType(ExcelChartType.ColumnClustered);
combo.AddSeries("C2:C10", "A2:A10", name: "Line").AsChartType(ExcelChartType.Line).OnSecondaryAxis();
combo.SetPosition(2, 13, 16, 20);

wb.Save(@"C:\temp\demo.xlsx");
```

### High-Volume Streaming (100k+ rows)

```csharp
using (var sw = ExcelStreamingWorkbook.Create(@"C:\temp\large.xlsx"))
{
    var big = sw.AddWorksheet("Rows");
    for (int r = 1; r <= 100_000; r++)
    {
        big.WriteRow(r, row =>
        {
            row.Text(1, "R" + r);
            row.Number(2, r);
        });
    }
}
```

---

## Quick Usage: Vector PDF Reports

PdfReport creates publication-ready A4 landscape reports. Series are populated from in-memory arrays rather than worksheet ranges.

```csharp
using Conq.Noble.Spreadox;

// Register optional custom TrueType font at runtime
// PdfFontRegistry.Register("Montserrat", @"C:\fonts\Montserrat-Regular.ttf");

using var report = PdfReport.Create(PdfPageSize.A4Landscape);
report.Margins = PdfMargins.Millimeters(10, 10, 10, 14);

// 1. Cover Page
report.Cover.Title = "Executive Analytics Report";
report.Cover.Company = "Conq.Noble Inc.";
report.Cover.Date = DateTime.Now;
report.Cover.Provider = "Spreadox Engine";
report.Cover.AddFact("Scope", "12 Months Performance");

// 2. Running Header (drawn on every page)
report.Header.Title = "Analyst Native Report";
report.Header.Text = "Confidential • Internal Distribution Only";
report.Header.FontSize = 11;
report.Header.Bold = true;
report.Header.TitleColor = ExcelColor.FromRgb(0, 82, 147);
// report.Header.Logo = PdfImage.FromFile("logo.png");

// 3. Running Footer (drawn on every page)
report.Footer.Text = "Spreadox Reporting Suite";
report.Footer.PageNumberFormat = "Page {0} / {1}";

// 4. Add Chart Pages
string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
double[] values = { 12.4, 11.8, 13.1, 12.0, 14.2, 13.6, 12.9, 15.0, 14.4, 13.2, 12.7, 13.8 };

var page = report.AddChartPage("Station Measurements");
page.Chart.Type = ExcelChartType.LineMarker;
page.Chart.AddSeries("Actual", months, values).SetColor(ExcelColor.FromRgb(0, 82, 147));

// Reference control lines
page.Chart.AddConstantLine("UCL", 15.5, ExcelColor.Red);
page.Chart.AddConstantLine("CL", 13.0, ExcelColor.Green);
page.Chart.AddConstantLine("LCL", 10.5, ExcelColor.Red);

// Right-hand summary table (up to 28 rows)
page.Summary.Add("Mean", "13.1");
page.Summary.Add("Max", "15.0");
page.Summary.Add("Min", "11.8");

report.Save(@"C:\temp\report.pdf");
```

---

## Runtime Custom Font Support

Spreadox embeds Liberation Sans by default. You can register any TrueType (.ttf) font at runtime before building reports:

```csharp
// From file:
PdfFontRegistry.Register("Roboto", @"C:\fonts\Roboto-Regular.ttf");

// Or from byte array (e.g. from file upload or database):
PdfFontRegistry.Register("CustomFont", fontBytes);

// Then specify it on Header, Footer, or chart elements:
report.Header.FontFamily = "Roboto";
report.Footer.FontFamily = "Roboto";
```

---

## DOM vs. Streaming Excel

| Feature | ExcelWorkbook | ExcelStreamingWorkbook |
|---|---|---|
| **Cell Model** | In-memory grid | Rows flushed as written |
| **Charts / Pictures / Conditional Format** | Full support | Not supported (raw data only) |
| **Large Exports** | Recommended up to tens of thousands of rows | Ideal for 100,000+ to millions of rows |
| **Strings Storage** | Shared string table | inlineStr (default) or SST |

---

## Performance Benchmarks & Quality Assurance

Spreadox is verified with an automated suite of **112 tests** across 7 quality categories (Memory, Concurrency, OpenXml Validation, Vector PDF, Charts, Boundaries, and NuGet Integrity):

- **1,000,000 Rows Export:** **22.1 seconds** (40 MB file, 667 MB peak working set).
- **100,000 Rows Export:** **2.1 seconds** (3.9 MB file, 321 MB peak working set).
- **Concurrency Isolation:** Verified across 1, 5, 10, 25, 50, and 100 simultaneous independent workbook exports.
- **OpenXML Schema Compliance:** 100% clean validation via `DocumentFormat.OpenXml.Validation.OpenXmlValidator` (0 errors).
- **Memory Leak Protection:** Verified zero memory leaks across repeated batch export cycles.

👉 **Full benchmark metrics, memory profiles, and compatibility tables:** [docs/TEST_RESULTS.md](docs/TEST_RESULTS.md)

---

## Interactive Samples

Explore the `samples/` directory:

1. **samples/ReportBuilderMvc (Live Web Demo):**
   - 🌐 **Live Demo on Vercel:** [https://spreadox-report-mvc.vercel.app/](https://spreadox-report-mvc.vercel.app/)
   - Full ASP.NET Core MVC Web Studio with **English / Turkish** localization.
   - Dynamic random data generator with customizable series count.
   - Per-series chart type picker (Column, Bar, Line, Area, Pie, Combo, etc.).
   - Live header/footer customization with logo file upload.
   - Web-based .ttf TrueType font uploader and instant registration.
   - Single-page vs. separate-page export options for both **Excel (.xlsx)** and **PDF (.pdf)**.
2. **samples/MeasurementChartSample:** Focused console demo for combo charts, statistics, and Excel styling.
3. **samples/PdfReportSample:** Demonstrates multi-page vector PDF reports with cover, combo charts, and control lines.
4. **samples/FullDemo.NetCore / FullDemo.Net48:** Deep dive showcasing every Spreadox API feature.

Run the MVC Studio sample:
```bash
cd samples/ReportBuilderMvc
dotnet run
```

---

## Packaging and Publishing to NuGet

### Building the NuGet Package locally:

```bash
dotnet pack src/Conq.Noble.Spreadox/Conq.Noble.Spreadox.csproj -c Release -o artifacts
```

This generates:
- `artifacts/Spreadox.1.0.0.nupkg`
- `artifacts/Spreadox.1.0.0.snupkg` (Symbols package)

### Publishing to NuGet.org:

Using the .NET CLI:
```bash
dotnet nuget push artifacts/Spreadox.1.0.0.nupkg --api-key <YOUR_NUGET_API_KEY> --source https://api.nuget.org/v3/index.json
```

Or via GitHub Actions:
Tagging a release as `v0.9.0` triggers the automated CI/CD publishing workflow when `NUGET_API_KEY` is configured in repository secrets.

---

## License

MIT — see [LICENSE](https://github.com/FatihBaycu/Spreadox/blob/main/LICENSE). Depends on DocumentFormat.OpenXml (MIT) and PdfSharpCore (MIT). Embedded Liberation Sans fonts are SIL Open Font License.
