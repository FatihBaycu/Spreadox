using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Conq.Noble.Spreadox.Internal
{
    /// <summary>
    /// Maps the public domain model to an Open XML spreadsheet package.
    /// </summary>
    internal static class OpenXmlWorkbookWriter
    {
        public static void Write(ExcelWorkbook workbook, Stream stream)
        {
            // Prefer writing directly into a seekable caller stream (avoids a full second package buffer).
            if (stream.CanSeek)
            {
                long start = stream.Position;
                WriteCore(workbook, stream);
                stream.Flush();
                // Leave position at end of written package (common for MemoryStream callers).
                if (stream.CanSeek && stream.Position == start)
                {
                    stream.Position = stream.Length;
                }

                return;
            }

            using (MemoryStream buffer = new MemoryStream())
            {
                WriteCore(workbook, buffer);
                buffer.Position = 0;
                buffer.CopyTo(stream);
                stream.Flush();
            }
        }

        private static void WriteCore(ExcelWorkbook workbook, Stream stream)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var styleCache = new StyleIndexCache();
                // Pre-register styles from all cells so stylesheet is complete before sheet write.
                foreach (ExcelWorksheet worksheet in workbook.Worksheets)
                {
                    foreach (ExcelCell cell in worksheet.EnumeratePersistableCells())
                    {
                        bool dateDefault = cell.Kind == CellValueKind.DateTime
                                           || (cell.Kind == CellValueKind.Formula && cell.Value is DateTime);
                        styleCache.GetStyleIndex(cell.StyleOrNull, dateDefault && (cell.StyleOrNull?.NumberFormat is null));
                    }
                }

                WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = BuildStylesheetWithConditionalFormats(styleCache, workbook);
                stylesPart.Stylesheet.Save();

                var sst = new SharedStringTableBuilder();
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                uint sheetId = 1;
                foreach (ExcelWorksheet worksheet in workbook.Worksheets)
                {
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    // Stream sheetData with OpenXmlWriter (DOM cells → SAX package XML).
                    WriteWorksheetPart(worksheetPart, worksheet, sst, styleCache);
                    // Reload root so chart attachment can append Drawing.
                    _ = worksheetPart.Worksheet;
                    ChartPartBuilder.AttachCharts(worksheetPart, worksheet);
                    worksheetPart.Worksheet!.Save();

                    sheets.AppendChild(new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = sheetId,
                        Name = worksheet.Name
                    });

                    sheetId++;
                }

                if (workbook.NamedRanges.Count > 0)
                {
                    var definedNames = new DefinedNames();
                    foreach (ExcelNamedRange named in workbook.NamedRanges)
                    {
                        definedNames.AppendChild(new DefinedName
                        {
                            Name = named.Name,
                            Text = named.RefersTo
                        });
                    }

                    workbookPart.Workbook.AppendChild(definedNames);
                }

                SharedStringTablePart sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringPart.SharedStringTable = BuildSharedStringTable(sst);
                sharedStringPart.SharedStringTable.Save();

                workbookPart.Workbook.Save();
            }
        }

        private static Stylesheet BuildStylesheetWithConditionalFormats(StyleIndexCache styleCache, ExcelWorkbook workbook)
        {
            Stylesheet stylesheet = styleCache.BuildStylesheet();

            bool anyCf = workbook.Worksheets.Any(w => w.ConditionalFormats.Count > 0);
            if (!anyCf)
            {
                return stylesheet;
            }

            var dxfs = new DifferentialFormats();
            uint count = 0;
            foreach (ExcelWorksheet worksheet in workbook.Worksheets)
            {
                foreach (ExcelConditionalFormat rule in worksheet.ConditionalFormats)
                {
                    if (rule.Kind != ExcelConditionalFormatKind.CellValue)
                    {
                        continue;
                    }

                    var pattern = new PatternFill { PatternType = PatternValues.Solid };
                    pattern.AppendChild(new ForegroundColor { Rgb = "FF" + rule.FillColor.ToRgbHex() });
                    pattern.AppendChild(new BackgroundColor { Rgb = "FF" + rule.FillColor.ToRgbHex() });
                    dxfs.AppendChild(new DifferentialFormat(new Fill(pattern)));
                    count++;
                }
            }

            if (count == 0)
            {
                return stylesheet;
            }

            dxfs.Count = count;
            stylesheet.AppendChild(dxfs);
            return stylesheet;
        }

        /// <summary>
        /// Writes worksheet XML with <see cref="OpenXmlWriter"/> so large <c>sheetData</c> is not held as an OpenXml DOM tree.
        /// Domain cells may still reside in <see cref="ExcelWorksheet"/> memory (use <see cref="ExcelStreamingWorkbook"/> to avoid that).
        /// </summary>
        private static void WriteWorksheetPart(
            WorksheetPart worksheetPart,
            ExcelWorksheet model,
            SharedStringTableBuilder sst,
            StyleIndexCache styleCache)
        {
            using (OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart))
            {
                writer.WriteStartElement(new Worksheet());

                if (model.FreezeRows > 0 || model.FreezeColumns > 0)
                {
                    writer.WriteElement(BuildSheetViews(model));
                }

                if (model.ColumnWidths.Count > 0)
                {
                    writer.WriteStartElement(new Columns());
                    foreach (KeyValuePair<int, double> pair in model.ColumnWidths.OrderBy(p => p.Key))
                    {
                        writer.WriteElement(new Column
                        {
                            Min = (uint)pair.Key,
                            Max = (uint)pair.Key,
                            Width = pair.Value,
                            CustomWidth = true
                        });
                    }

                    writer.WriteEndElement(); // Columns
                }

                writer.WriteStartElement(new SheetData());

                IEnumerable<IGrouping<int, ExcelCell>> rows = model
                    .EnumeratePersistableCells()
                    .GroupBy(c => c.Row)
                    .OrderBy(g => g.Key);

                HashSet<int> rowsWithHeight = new HashSet<int>(model.RowHeights.Keys);

                foreach (IGrouping<int, ExcelCell> rowGroup in rows)
                {
                    writer.WriteStartElement(CreateRow(rowGroup.Key, model));
                    rowsWithHeight.Remove(rowGroup.Key);
                    foreach (ExcelCell cell in rowGroup.OrderBy(c => c.Column))
                    {
                        writer.WriteElement(CreateCell(cell, sst, styleCache));
                    }

                    writer.WriteEndElement(); // Row
                }

                foreach (int rowIndex in rowsWithHeight.OrderBy(i => i))
                {
                    writer.WriteElement(CreateRow(rowIndex, model));
                }

                writer.WriteEndElement(); // SheetData

                if (!string.IsNullOrEmpty(model.AutoFilterRange))
                {
                    writer.WriteElement(new AutoFilter { Reference = model.AutoFilterRange });
                }

                if (model.MergedRanges.Count > 0)
                {
                    writer.WriteStartElement(new MergeCells { Count = (uint)model.MergedRanges.Count });
                    foreach (ExcelRange range in model.MergedRanges)
                    {
                        writer.WriteElement(new MergeCell { Reference = range.A1 });
                    }

                    writer.WriteEndElement(); // MergeCells
                }

                if (model.ConditionalFormats.Count > 0)
                {
                    WriteConditionalFormatting(writer, model);
                }

                writer.WriteEndElement(); // Worksheet
            }
        }

        private static void WriteConditionalFormatting(OpenXmlWriter writer, ExcelWorksheet model)
        {
            int dxfIndex = 0;
            foreach (ExcelWorksheet sheet in model.Workbook.Worksheets)
            {
                if (!ReferenceEquals(sheet, model))
                {
                    foreach (ExcelConditionalFormat prior in sheet.ConditionalFormats)
                    {
                        if (prior.Kind == ExcelConditionalFormatKind.CellValue)
                        {
                            dxfIndex++;
                        }
                    }

                    continue;
                }

                int priority = 1;
                foreach (ExcelConditionalFormat rule in model.ConditionalFormats)
                {
                    var cf = new ConditionalFormatting
                    {
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = rule.Range.A1 }
                    };

                    if (rule.Kind == ExcelConditionalFormatKind.ColorScale && rule.ColorScale != null)
                    {
                        var cfRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.ColorScale,
                            Priority = rule.Priority > 0 ? rule.Priority : priority
                        };
                        cfRule.AppendChild(BuildColorScale(rule.ColorScale));
                        cf.AppendChild(cfRule);
                    }
                    else
                    {
                        var cfRule = new ConditionalFormattingRule
                        {
                            Type = ConditionalFormatValues.CellIs,
                            Operator = ToCfOperator(rule.Operator),
                            Priority = rule.Priority > 0 ? rule.Priority : priority,
                            FormatId = (uint)dxfIndex
                        };
                        cfRule.AppendChild(new Formula(rule.FormulaOperand));
                        cf.AppendChild(cfRule);
                        dxfIndex++;
                    }

                    writer.WriteElement(cf);
                    priority++;
                }

                break;
            }
        }

        private static ColorScale BuildColorScale(ExcelColorScale scale)
        {
            var node = new ColorScale();
            AppendCfvo(node, scale.Low);
            if (scale.Midpoint != null)
            {
                AppendCfvo(node, scale.Midpoint);
            }

            AppendCfvo(node, scale.High);
            node.AppendChild(ToSpreadsheetColor(scale.Low.Color));
            if (scale.Midpoint != null)
            {
                node.AppendChild(ToSpreadsheetColor(scale.Midpoint.Color));
            }

            node.AppendChild(ToSpreadsheetColor(scale.High.Color));
            return node;
        }

        private static void AppendCfvo(ColorScale node, ExcelColorScaleStop stop)
        {
            var cfvo = new ConditionalFormatValueObject { Type = ToCfvoType(stop.Type) };
            if (stop.Type != ExcelColorScaleValueType.Min && stop.Type != ExcelColorScaleValueType.Max)
            {
                cfvo.Val = (stop.Value ?? 0).ToString(CultureInfo.InvariantCulture);
            }

            node.AppendChild(cfvo);
        }

        private static ConditionalFormatValueObjectValues ToCfvoType(ExcelColorScaleValueType type)
        {
            switch (type)
            {
                case ExcelColorScaleValueType.Percent: return ConditionalFormatValueObjectValues.Percent;
                case ExcelColorScaleValueType.Min: return ConditionalFormatValueObjectValues.Min;
                case ExcelColorScaleValueType.Max: return ConditionalFormatValueObjectValues.Max;
                case ExcelColorScaleValueType.Number: return ConditionalFormatValueObjectValues.Number;
                default: return ConditionalFormatValueObjectValues.Percentile;
            }
        }

        private static Color ToSpreadsheetColor(ExcelColor color) =>
            new Color { Rgb = "FF" + color.ToRgbHex() };

        private static ConditionalFormattingOperatorValues ToCfOperator(ConditionalOperator op)
        {
            switch (op)
            {
                case ConditionalOperator.GreaterThan: return ConditionalFormattingOperatorValues.GreaterThan;
                case ConditionalOperator.GreaterThanOrEqual: return ConditionalFormattingOperatorValues.GreaterThanOrEqual;
                case ConditionalOperator.LessThan: return ConditionalFormattingOperatorValues.LessThan;
                case ConditionalOperator.LessThanOrEqual: return ConditionalFormattingOperatorValues.LessThanOrEqual;
                case ConditionalOperator.Equal: return ConditionalFormattingOperatorValues.Equal;
                case ConditionalOperator.NotEqual: return ConditionalFormattingOperatorValues.NotEqual;
                default: return ConditionalFormattingOperatorValues.GreaterThan;
            }
        }

        private static SheetViews BuildSheetViews(ExcelWorksheet model)
        {
            var view = new SheetView { WorkbookViewId = 0 };
            if (model.FreezeRows > 0 || model.FreezeColumns > 0)
            {
                view.AppendChild(new Pane
                {
                    VerticalSplit = model.FreezeRows,
                    HorizontalSplit = model.FreezeColumns,
                    TopLeftCell = CellAddress.Format(model.FreezeRows + 1, model.FreezeColumns + 1),
                    ActivePane = model.FreezeRows > 0 && model.FreezeColumns > 0
                        ? PaneValues.BottomRight
                        : model.FreezeRows > 0 ? PaneValues.BottomLeft : PaneValues.TopRight,
                    State = PaneStateValues.Frozen
                });
            }

            return new SheetViews(view);
        }

        private static Row CreateRow(int rowIndex, ExcelWorksheet model)
        {
            var row = new Row { RowIndex = (uint)rowIndex };
            if (model.TryGetRowHeight(rowIndex, out double height))
            {
                row.Height = height;
                row.CustomHeight = true;
            }

            return row;
        }

        private static Cell CreateCell(ExcelCell model, SharedStringTableBuilder sst, StyleIndexCache styleCache)
        {
            Cell cell = new Cell
            {
                CellReference = model.Address
            };

            bool dateDefault = model.Kind == CellValueKind.DateTime
                               || (model.HasFormula && model.Value is DateTime);
            bool useDateDefault = dateDefault && string.IsNullOrEmpty(model.StyleOrNull?.NumberFormat);
            uint styleIndex = styleCache.GetStyleIndex(model.StyleOrNull, useDateDefault);
            if (styleIndex != 0 || model.HasStyle || useDateDefault)
            {
                cell.StyleIndex = styleIndex;
            }

            if (model.HasFormula)
            {
                cell.CellFormula = new CellFormula(model.Formula);
                WriteCachedValue(cell, model);
                return cell;
            }

            switch (model.Kind)
            {
                case CellValueKind.String:
                    {
                        string text = model.GetString() ?? string.Empty;
                        int index = sst.GetOrAdd(text);
                        cell.DataType = CellValues.SharedString;
                        cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
                        break;
                    }
                case CellValueKind.Integer:
                    cell.CellValue = new CellValue(model.GetInt()!.Value.ToString(CultureInfo.InvariantCulture));
                    break;
                case CellValueKind.Double:
                    cell.CellValue = new CellValue(model.GetDouble()!.Value.ToString("G17", CultureInfo.InvariantCulture));
                    break;
                case CellValueKind.Decimal:
                    cell.CellValue = new CellValue(model.GetDecimal()!.Value.ToString(CultureInfo.InvariantCulture));
                    break;
                case CellValueKind.Boolean:
                    cell.DataType = CellValues.Boolean;
                    cell.CellValue = new CellValue(model.GetBoolean()!.Value ? "1" : "0");
                    break;
                case CellValueKind.DateTime:
                    {
                        DateTime dt = model.GetDateTime()!.Value;
                        cell.CellValue = new CellValue(ToExcelSerialDate(dt).ToString("G17", CultureInfo.InvariantCulture));
                        break;
                    }
                case CellValueKind.Empty:
                    // style-only cell
                    break;
                default:
                    throw new InvalidOperationException($"Cannot serialize cell kind '{model.Kind}'.");
            }

            return cell;
        }

        private static void WriteCachedValue(Cell cell, ExcelCell model)
        {
            if (model.Value is null)
            {
                return;
            }

            switch (model.Value)
            {
                case string s:
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(s);
                    break;
                case int i:
                    cell.CellValue = new CellValue(i.ToString(CultureInfo.InvariantCulture));
                    break;
                case double d:
                    cell.CellValue = new CellValue(d.ToString("G17", CultureInfo.InvariantCulture));
                    break;
                case decimal m:
                    cell.CellValue = new CellValue(m.ToString(CultureInfo.InvariantCulture));
                    break;
                case bool b:
                    cell.DataType = CellValues.Boolean;
                    cell.CellValue = new CellValue(b ? "1" : "0");
                    break;
                case DateTime dt:
                    cell.CellValue = new CellValue(ToExcelSerialDate(dt).ToString("G17", CultureInfo.InvariantCulture));
                    break;
            }
        }

        private static SharedStringTable BuildSharedStringTable(SharedStringTableBuilder sst)
        {
            SharedStringTable table = new SharedStringTable
            {
                Count = (uint)sst.Values.Count,
                UniqueCount = (uint)sst.Values.Count
            };

            foreach (string value in sst.Values)
            {
                table.AppendChild(new SharedStringItem(new Text(SanitizeXmlString(value))));
            }

            return table;
        }

        internal static string SanitizeXmlString(string? text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            bool hasInvalid = false;
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r')
                {
                    hasInvalid = true;
                    break;
                }
            }

            if (!hasInvalid)
            {
                return text;
            }

            var sb = new System.Text.StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch >= 0x20 || ch == '\t' || ch == '\n' || ch == '\r')
                {
                    sb.Append(ch);
                }
            }

            return sb.ToString();
        }

        internal static double ToExcelSerialDate(DateTime value) => value.ToOADate();
    }
}
