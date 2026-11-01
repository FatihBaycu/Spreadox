using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Conq.Noble.Spreadox.Internal
{
    /// <summary>
    /// Loads an Open XML package into the public domain model.
    /// </summary>
    internal static class OpenXmlWorkbookReader
    {
        public static void Load(ExcelWorkbook workbook, Stream stream)
        {
            using (MemoryStream buffer = CopyToMemory(stream))
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(buffer, false))
            {
                WorkbookPart? workbookPart = document.WorkbookPart;
                if (workbookPart?.Workbook?.Sheets is null)
                {
                    throw new InvalidDataException("The package does not contain a valid Excel workbook or worksheets.");
                }

                IList<string> sharedStrings = LoadSharedStrings(workbookPart);
                StylesheetInfo styles = LoadStyles(workbookPart);

                workbook.ClearWorksheetsForLoad();
                workbook.ClearNamedRangesForLoad();

                foreach (Sheet sheet in workbookPart.Workbook.Sheets.Elements<Sheet>())
                {
                    string name = sheet.Name?.Value ?? $"Sheet{workbook.Worksheets.Count + 1}";
                    ExcelWorksheet worksheet = new ExcelWorksheet(workbook, name);
                    workbook.AddLoadedWorksheet(worksheet);

                    if (sheet.Id?.Value is null)
                    {
                        continue;
                    }

                    if (!(workbookPart.GetPartById(sheet.Id.Value) is WorksheetPart worksheetPart))
                    {
                        continue;
                    }

                    LoadWorksheetExtras(worksheet, worksheetPart);
                    LoadCells(worksheet, worksheetPart, sharedStrings, styles);
                    ChartPartReader.LoadCharts(worksheet, worksheetPart);
                }

                DefinedNames? definedNames = workbookPart.Workbook.DefinedNames;
                if (definedNames != null)
                {
                    foreach (DefinedName defined in definedNames.Elements<DefinedName>())
                    {
                        if (string.IsNullOrWhiteSpace(defined.Name?.Value) || string.IsNullOrWhiteSpace(defined.Text))
                        {
                            continue;
                        }

                        workbook.AddLoadedNamedRange(new ExcelNamedRange(defined.Name.Value, defined.Text));
                    }
                }
            }
        }

        private static MemoryStream CopyToMemory(Stream stream)
        {
            MemoryStream buffer = new MemoryStream();
            if (stream.CanSeek)
            {
                long original = stream.Position;
                stream.CopyTo(buffer);
                stream.Position = original;
            }
            else
            {
                stream.CopyTo(buffer);
            }

            buffer.Position = 0;
            return buffer;
        }

        private static void LoadWorksheetExtras(ExcelWorksheet worksheet, WorksheetPart worksheetPart)
        {
            Worksheet? ws = worksheetPart.Worksheet;
            if (ws is null)
            {
                return;
            }

            SheetViews? views = ws.GetFirstChild<SheetViews>();
            Pane? pane = views?.GetFirstChild<SheetView>()?.GetFirstChild<Pane>();
            if (pane != null)
            {
                int freezeRows = (int)(pane.VerticalSplit?.Value ?? 0);
                int freezeCols = (int)(pane.HorizontalSplit?.Value ?? 0);
                worksheet.LoadFreezePanes(freezeRows, freezeCols);
            }

            Columns? columns = ws.GetFirstChild<Columns>();
            if (columns != null)
            {
                foreach (Column col in columns.Elements<Column>())
                {
                    uint min = col.Min?.Value ?? 1;
                    uint max = col.Max?.Value ?? min;
                    double width = col.Width?.Value ?? 8.43;
                    for (uint c = min; c <= max; c++)
                    {
                        worksheet.LoadColumnWidth((int)c, width);
                    }
                }
            }

            AutoFilter? autoFilter = ws.GetFirstChild<AutoFilter>();
            if (autoFilter?.Reference?.Value != null)
            {
                worksheet.LoadAutoFilter(autoFilter.Reference.Value);
            }

            MergeCells? merges = ws.GetFirstChild<MergeCells>();
            if (merges != null)
            {
                foreach (MergeCell merge in merges.Elements<MergeCell>())
                {
                    if (!string.IsNullOrEmpty(merge.Reference?.Value))
                    {
                        worksheet.LoadMergedRange(ExcelRange.Parse(merge.Reference.Value));
                    }
                }
            }

            foreach (ConditionalFormatting cf in ws.Elements<ConditionalFormatting>())
            {
                string? rangeText = cf.SequenceOfReferences?.InnerText
                                    ?? cf.SequenceOfReferences?.Items?.FirstOrDefault()?.Value;
                ConditionalFormattingRule? rule = cf.GetFirstChild<ConditionalFormattingRule>();
                if (string.IsNullOrEmpty(rangeText) || rule is null)
                {
                    continue;
                }

                if (rule.Type?.Value == ConditionalFormatValues.ColorScale)
                {
                    ExcelConditionalFormat? scale = TryReadColorScale(rangeText!, rule);
                    if (scale != null)
                    {
                        worksheet.LoadConditionalFormat(scale);
                    }

                    continue;
                }

                ConditionalOperator op = FromCfOperator(rule.Operator?.Value);
                string operand = rule.GetFirstChild<Formula>()?.Text ?? "0";
                ExcelColor fill = ExcelColor.FromRgb(255, 199, 206);
                worksheet.LoadConditionalFormat(new ExcelConditionalFormat(rangeText, op, operand, fill));
            }
        }

        private static ExcelConditionalFormat? TryReadColorScale(string rangeText, ConditionalFormattingRule rule)
        {
            ColorScale? scale = rule.GetFirstChild<ColorScale>();
            if (scale is null)
            {
                return null;
            }

            List<ConditionalFormatValueObject> vos = scale.Elements<ConditionalFormatValueObject>().ToList();
            List<Color> colors = scale.Elements<Color>().ToList();
            if (vos.Count < 2 || colors.Count < 2)
            {
                return null;
            }

            ExcelColorScaleStop low = ReadStop(vos[0], colors[0]);
            ExcelColorScaleStop high = ReadStop(vos[vos.Count - 1], colors[colors.Count - 1]);
            ExcelColorScaleStop? mid = vos.Count >= 3 && colors.Count >= 3
                ? ReadStop(vos[1], colors[1])
                : null;
            return ExcelConditionalFormat.ColorScalePercentile(
                rangeText,
                low.Color,
                mid?.Color ?? ExcelColor.FromHex("FFD166"),
                high.Color);
        }

        private static ExcelColorScaleStop ReadStop(ConditionalFormatValueObject vo, Color color)
        {
            ExcelColorScaleValueType type = ExcelColorScaleValueType.Percentile;
            if (vo.Type?.Value == ConditionalFormatValueObjectValues.Percent) type = ExcelColorScaleValueType.Percent;
            else if (vo.Type?.Value == ConditionalFormatValueObjectValues.Min) type = ExcelColorScaleValueType.Min;
            else if (vo.Type?.Value == ConditionalFormatValueObjectValues.Max) type = ExcelColorScaleValueType.Max;
            else if (vo.Type?.Value == ConditionalFormatValueObjectValues.Number) type = ExcelColorScaleValueType.Number;

            double? value = null;
            if (!string.IsNullOrEmpty(vo.Val?.Value)
                && double.TryParse(vo.Val.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
            {
                value = parsed;
            }

            string? rgbText = color.Rgb?.Value;
            ExcelColor rgb = !string.IsNullOrEmpty(rgbText) ? (ParseRgb(rgbText!) ?? ExcelColor.Gray) : ExcelColor.Gray;
            return new ExcelColorScaleStop(type, value, rgb);
        }

        private static ConditionalOperator FromCfOperator(ConditionalFormattingOperatorValues? value)
        {
            if (value == null) return ConditionalOperator.GreaterThan;
            if (value == ConditionalFormattingOperatorValues.GreaterThanOrEqual) return ConditionalOperator.GreaterThanOrEqual;
            if (value == ConditionalFormattingOperatorValues.LessThan) return ConditionalOperator.LessThan;
            if (value == ConditionalFormattingOperatorValues.LessThanOrEqual) return ConditionalOperator.LessThanOrEqual;
            if (value == ConditionalFormattingOperatorValues.Equal) return ConditionalOperator.Equal;
            if (value == ConditionalFormattingOperatorValues.NotEqual) return ConditionalOperator.NotEqual;
            return ConditionalOperator.GreaterThan;
        }

        private static void LoadCells(
            ExcelWorksheet worksheet,
            WorksheetPart worksheetPart,
            IList<string> sharedStrings,
            StylesheetInfo styles)
        {
            SheetData? sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
            if (sheetData is null)
            {
                return;
            }

            foreach (Row row in sheetData.Elements<Row>())
            {
                if (row.RowIndex?.Value != null && row.CustomHeight?.Value == true && row.Height?.Value != null)
                {
                    worksheet.LoadRowHeight((int)row.RowIndex.Value, row.Height.Value);
                }

                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.IsNullOrEmpty(cell.CellReference?.Value))
                    {
                        continue;
                    }

                    CellAddress.Parse(cell.CellReference.Value, out int r, out int c);
                    ExcelCell excelCell = worksheet[r, c];
                    ApplyCellValue(excelCell, cell, sharedStrings, styles);
                }
            }
        }

        private static void ApplyCellValue(
            ExcelCell excelCell,
            Cell cell,
            IList<string> sharedStrings,
            StylesheetInfo styles)
        {
            ExcelCellStyle? style = null;
            uint styleIndex = cell.StyleIndex?.Value ?? 0;
            if (styles.TryGetStyle(styleIndex, out ExcelCellStyle? resolved))
            {
                style = resolved;
            }

            string? formula = cell.CellFormula?.Text;
            if (!string.IsNullOrEmpty(formula))
            {
                object? cached = ReadCachedValue(cell, sharedStrings, styles.IsDateStyle(styleIndex));
                excelCell.Load(CellValueKind.Formula, cached, formula, style);
                return;
            }

            string? raw = cell.CellValue?.InnerText;
            CellValues? dataType = cell.DataType?.Value;

            if (dataType == CellValues.SharedString)
            {
                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)
                    && index >= 0 && index < sharedStrings.Count)
                {
                    excelCell.Load(CellValueKind.String, sharedStrings[index], null, style);
                }
                else
                {
                    excelCell.Load(CellValueKind.Empty, null, null, style);
                }

                return;
            }

            if (dataType == CellValues.InlineString)
            {
                string text = cell.InlineString?.Text?.Text
                              ?? cell.InlineString?.InnerText
                              ?? raw
                              ?? string.Empty;
                excelCell.Load(CellValueKind.String, text, null, style);
                return;
            }

            if (dataType == CellValues.Boolean)
            {
                bool value = raw == "1" || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);
                excelCell.Load(CellValueKind.Boolean, value, null, style);
                return;
            }

            if (dataType == CellValues.String || dataType == CellValues.Error)
            {
                excelCell.Load(CellValueKind.String, raw ?? string.Empty, null, style);
                return;
            }

            if (string.IsNullOrEmpty(raw))
            {
                excelCell.Load(CellValueKind.Empty, null, null, style);
                return;
            }

            if (styles.IsDateStyle(styleIndex)
                && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double oaDate))
            {
                try
                {
                    excelCell.Load(CellValueKind.DateTime, DateTime.FromOADate(oaDate), null, style);
                    return;
                }
                catch (ArgumentException)
                {
                    // Fall through.
                }
            }

            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            {
                if (number == Math.Truncate(number)
                    && number >= int.MinValue && number <= int.MaxValue
                    && raw.IndexOf('.') < 0 && raw.IndexOf('e') < 0 && raw.IndexOf('E') < 0)
                {
                    excelCell.Load(CellValueKind.Integer, (int)number, null, style);
                }
                else
                {
                    excelCell.Load(CellValueKind.Double, number, null, style);
                }

                return;
            }

            excelCell.Load(CellValueKind.String, raw, null, style);
        }

        private static object? ReadCachedValue(Cell cell, IList<string> sharedStrings, bool treatAsDate)
        {
            string? raw = cell.CellValue?.InnerText;
            if (string.IsNullOrEmpty(raw))
            {
                return null;
            }

            CellValues? dataType = cell.DataType?.Value;
            if (dataType == CellValues.SharedString
                && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)
                && index >= 0 && index < sharedStrings.Count)
            {
                return sharedStrings[index];
            }

            if (dataType == CellValues.Boolean)
            {
                return raw == "1" || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);
            }

            if (dataType == CellValues.String || dataType == CellValues.InlineString)
            {
                return raw;
            }

            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            {
                if (treatAsDate)
                {
                    try { return DateTime.FromOADate(number); }
                    catch (ArgumentException) { /* ignore */ }
                }

                if (number == Math.Truncate(number)
                    && number >= int.MinValue && number <= int.MaxValue
                    && raw.IndexOf('.') < 0)
                {
                    return (int)number;
                }

                return number;
            }

            return raw;
        }

        private static IList<string> LoadSharedStrings(WorkbookPart workbookPart)
        {
            SharedStringTablePart? part = workbookPart.SharedStringTablePart;
            if (part?.SharedStringTable is null)
            {
                return Array.Empty<string>();
            }

            List<string> values = new List<string>();
            foreach (SharedStringItem item in part.SharedStringTable.Elements<SharedStringItem>())
            {
                values.Add(item.Text?.Text ?? item.InnerText ?? string.Empty);
            }

            return values;
        }

        private static StylesheetInfo LoadStyles(WorkbookPart workbookPart)
        {
            var info = new StylesheetInfo();
            Stylesheet? stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
            if (stylesheet?.CellFormats is null)
            {
                return info;
            }

            HashSet<uint> dateNumFmts = new HashSet<uint>
            {
                14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
            };

            Dictionary<uint, string> customFormats = new Dictionary<uint, string>();
            if (stylesheet.NumberingFormats != null)
            {
                foreach (NumberingFormat nf in stylesheet.NumberingFormats.Elements<NumberingFormat>())
                {
                    string? code = nf.FormatCode?.Value;
                    if (nf.NumberFormatId?.Value == null || code == null)
                    {
                        continue;
                    }

                    customFormats[nf.NumberFormatId.Value] = code;
                    if (LooksLikeDateFormat(code))
                    {
                        dateNumFmts.Add(nf.NumberFormatId.Value);
                    }
                }
            }

            List<ExcelFont> fonts = LoadFonts(stylesheet);
            List<ExcelFill> fills = LoadFills(stylesheet);
            List<ExcelBorders> borders = LoadBorders(stylesheet);

            uint index = 0;
            foreach (CellFormat format in stylesheet.CellFormats.Elements<CellFormat>())
            {
                uint numFmtId = format.NumberFormatId?.Value ?? 0;
                bool isDate = dateNumFmts.Contains(numFmtId);
                if (isDate)
                {
                    info.DateStyleIndexes.Add(index);
                }

                var style = new ExcelCellStyle();
                if (customFormats.TryGetValue(numFmtId, out string? custom))
                {
                    style.NumberFormat = custom;
                }
                else if (numFmtId == 1) style.NumberFormat = NumberFormats.Integer;
                else if (numFmtId == 2) style.NumberFormat = NumberFormats.TwoDecimals;
                else if (numFmtId == 3) style.NumberFormat = NumberFormats.Thousands;
                else if (numFmtId == 4) style.NumberFormat = NumberFormats.ThousandsTwoDecimals;
                else if (numFmtId == 9) style.NumberFormat = NumberFormats.Percent;
                else if (numFmtId == 10) style.NumberFormat = NumberFormats.PercentTwoDecimals;

                uint fontId = format.FontId?.Value ?? 0;
                if (fontId < fonts.Count)
                {
                    ExcelFont src = fonts[(int)fontId];
                    style.Font.Name = src.Name;
                    style.Font.Size = src.Size;
                    style.Font.Bold = src.Bold;
                    style.Font.Italic = src.Italic;
                    style.Font.Underline = src.Underline;
                    style.Font.Strikeout = src.Strikeout;
                    style.Font.Color = src.Color;
                }

                uint fillId = format.FillId?.Value ?? 0;
                if (fillId < fills.Count)
                {
                    ExcelFill src = fills[(int)fillId];
                    style.Fill.Pattern = src.Pattern;
                    style.Fill.ForegroundColor = src.ForegroundColor;
                    style.Fill.BackgroundColor = src.BackgroundColor;
                }

                uint borderId = format.BorderId?.Value ?? 0;
                if (borderId < borders.Count)
                {
                    ExcelBorders src = borders[(int)borderId];
                    style.Borders.Left.Style = src.Left.Style;
                    style.Borders.Left.Color = src.Left.Color;
                    style.Borders.Right.Style = src.Right.Style;
                    style.Borders.Right.Color = src.Right.Color;
                    style.Borders.Top.Style = src.Top.Style;
                    style.Borders.Top.Color = src.Top.Color;
                    style.Borders.Bottom.Style = src.Bottom.Style;
                    style.Borders.Bottom.Color = src.Bottom.Color;
                }

                if (format.Alignment != null)
                {
                    style.Alignment.Horizontal = FromHorizontal(format.Alignment.Horizontal?.Value);
                    style.Alignment.Vertical = FromVertical(format.Alignment.Vertical?.Value);
                    style.Alignment.WrapText = format.Alignment.WrapText?.Value ?? false;
                    style.Alignment.Indent = (int)(format.Alignment.Indent?.Value ?? 0);
                }

                info.Styles[index] = style;
                index++;
            }

            return info;
        }

        private static List<ExcelFont> LoadFonts(Stylesheet stylesheet)
        {
            var list = new List<ExcelFont>();
            Fonts? fonts = stylesheet.Fonts;
            if (fonts is null)
            {
                list.Add(new ExcelFont());
                return list;
            }

            foreach (Font font in fonts.Elements<Font>())
            {
                var model = new ExcelFont
                {
                    Bold = font.Bold != null,
                    Italic = font.Italic != null,
                    Underline = font.Underline != null,
                    Strikeout = font.Strike != null,
                    Size = font.FontSize?.Val?.HasValue == true ? font.FontSize.Val.Value : (double?)null,
                    Name = font.FontName?.Val?.Value
                };

                string? rgb = font.Color?.Rgb?.Value;
                if (!string.IsNullOrEmpty(rgb))
                {
                    model.Color = ParseRgb(rgb!);
                }

                list.Add(model);
            }

            return list;
        }

        private static List<ExcelFill> LoadFills(Stylesheet stylesheet)
        {
            var list = new List<ExcelFill>();
            Fills? fills = stylesheet.Fills;
            if (fills is null)
            {
                list.Add(new ExcelFill());
                return list;
            }

            foreach (Fill fill in fills.Elements<Fill>())
            {
                PatternFill? pattern = fill.PatternFill;
                var model = new ExcelFill();
                if (pattern?.PatternType?.Value == PatternValues.Solid)
                {
                    model.Pattern = ExcelFillPattern.Solid;
                    string? rgb = pattern.ForegroundColor?.Rgb ?? pattern.BackgroundColor?.Rgb;
                    if (!string.IsNullOrEmpty(rgb))
                    {
                        model.ForegroundColor = ParseRgb(rgb);
                        model.BackgroundColor = model.ForegroundColor;
                    }
                }

                list.Add(model);
            }

            return list;
        }

        private static List<ExcelBorders> LoadBorders(Stylesheet stylesheet)
        {
            var list = new List<ExcelBorders>();
            Borders? borders = stylesheet.Borders;
            if (borders is null)
            {
                list.Add(new ExcelBorders());
                return list;
            }

            foreach (Border border in borders.Elements<Border>())
            {
                var model = new ExcelBorders();
                ApplyEdge(model.Left, border.LeftBorder);
                ApplyEdge(model.Right, border.RightBorder);
                ApplyEdge(model.Top, border.TopBorder);
                ApplyEdge(model.Bottom, border.BottomBorder);
                list.Add(model);
            }

            return list;
        }

        private static void ApplyEdge(ExcelBorderEdge edge, LeftBorder? border)
        {
            ApplyEdgeCore(edge, border?.Style?.Value, border?.GetFirstChild<Color>()?.Rgb);
        }

        private static void ApplyEdge(ExcelBorderEdge edge, RightBorder? border)
        {
            ApplyEdgeCore(edge, border?.Style?.Value, border?.GetFirstChild<Color>()?.Rgb);
        }

        private static void ApplyEdge(ExcelBorderEdge edge, TopBorder? border)
        {
            ApplyEdgeCore(edge, border?.Style?.Value, border?.GetFirstChild<Color>()?.Rgb);
        }

        private static void ApplyEdge(ExcelBorderEdge edge, BottomBorder? border)
        {
            ApplyEdgeCore(edge, border?.Style?.Value, border?.GetFirstChild<Color>()?.Rgb);
        }

        private static void ApplyEdgeCore(ExcelBorderEdge edge, BorderStyleValues? style, string? rgb)
        {
            if (style == null)
            {
                return;
            }

            edge.Style = FromBorderStyle(style.Value);
            if (!string.IsNullOrEmpty(rgb))
            {
                edge.Color = ParseRgb(rgb!);
            }
        }

        private static ExcelBorderStyle FromBorderStyle(BorderStyleValues style)
        {
            if (style == BorderStyleValues.Thin) return ExcelBorderStyle.Thin;
            if (style == BorderStyleValues.Medium) return ExcelBorderStyle.Medium;
            if (style == BorderStyleValues.Thick) return ExcelBorderStyle.Thick;
            if (style == BorderStyleValues.Dashed) return ExcelBorderStyle.Dashed;
            if (style == BorderStyleValues.Dotted) return ExcelBorderStyle.Dotted;
            if (style == BorderStyleValues.Double) return ExcelBorderStyle.Double;
            if (style == BorderStyleValues.Hair) return ExcelBorderStyle.Hair;
            return ExcelBorderStyle.None;
        }

        private static ExcelHorizontalAlignment FromHorizontal(HorizontalAlignmentValues? value)
        {
            if (value == null) return ExcelHorizontalAlignment.General;
            if (value == HorizontalAlignmentValues.Left) return ExcelHorizontalAlignment.Left;
            if (value == HorizontalAlignmentValues.Center) return ExcelHorizontalAlignment.Center;
            if (value == HorizontalAlignmentValues.Right) return ExcelHorizontalAlignment.Right;
            if (value == HorizontalAlignmentValues.Fill) return ExcelHorizontalAlignment.Fill;
            if (value == HorizontalAlignmentValues.Justify) return ExcelHorizontalAlignment.Justify;
            return ExcelHorizontalAlignment.General;
        }

        private static ExcelVerticalAlignment FromVertical(VerticalAlignmentValues? value)
        {
            if (value == null) return ExcelVerticalAlignment.Bottom;
            if (value == VerticalAlignmentValues.Top) return ExcelVerticalAlignment.Top;
            if (value == VerticalAlignmentValues.Center) return ExcelVerticalAlignment.Center;
            if (value == VerticalAlignmentValues.Justify) return ExcelVerticalAlignment.Justify;
            return ExcelVerticalAlignment.Bottom;
        }

        private static ExcelColor? ParseRgb(string rgb)
        {
            try
            {
                if (rgb.Length == 8)
                {
                    return ExcelColor.FromHex(rgb);
                }

                if (rgb.Length == 6)
                {
                    return ExcelColor.FromHex(rgb);
                }
            }
            catch (ArgumentException)
            {
                return null;
            }

            return null;
        }

        private static bool LooksLikeDateFormat(string formatCode)
        {
            string lower = formatCode.ToLowerInvariant();
            if (lower.IndexOf("red", StringComparison.Ordinal) >= 0
                && lower.IndexOf('y') < 0
                && lower.IndexOf('d') < 0
                && lower.IndexOf('h') < 0
                && lower.IndexOf('m') < 0)
            {
                return false;
            }

            return lower.IndexOf('y') >= 0
                   || lower.IndexOf('d') >= 0
                   || lower.IndexOf('h') >= 0
                   || lower.IndexOf('s') >= 0
                   || (lower.IndexOf('m') >= 0 && (lower.IndexOf('y') >= 0 || lower.IndexOf('d') >= 0 || lower.IndexOf(':') >= 0));
        }

        private sealed class StylesheetInfo
        {
            public HashSet<uint> DateStyleIndexes { get; } = new HashSet<uint>();
            public Dictionary<uint, ExcelCellStyle> Styles { get; } = new Dictionary<uint, ExcelCellStyle>();

            public bool IsDateStyle(uint index) => DateStyleIndexes.Contains(index);

            public bool TryGetStyle(uint index, out ExcelCellStyle? style)
            {
                if (Styles.TryGetValue(index, out ExcelCellStyle? found) && found != null && !found.IsDefault)
                {
                    style = found.Clone();
                    return true;
                }

                style = null;
                return false;
            }
        }
    }
}
