using System;
using System.Collections.Generic;
using Conq.Noble.Spreadox.Internal;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// One worksheet being written via the streaming path. Call <see cref="Complete"/> (or dispose the workbook)
    /// after the last row. Column widths must be set before the first WriteRow call.
    /// </summary>
    public sealed class ExcelStreamingWorksheet
    {
        private readonly StreamingPackageWriter _package;
        private readonly StreamingSheetWriter _sheet;
        private readonly ExcelStreamingRowWriter _rowWriter;
        private bool _completed;

        internal ExcelStreamingWorksheet(StreamingPackageWriter package, StreamingSheetWriter sheet)
        {
            _package = package;
            _sheet = sheet;
            _rowWriter = new ExcelStreamingRowWriter(package, sheet, 0);
        }

        public string Name => _sheet.Name;

        public void SetColumnWidth(string columnLetter, double width)
        {
            ThrowIfCompleted();
            CellAddress.Parse(columnLetter + "1", out _, out int column);
            _sheet.SetColumnWidth(column, width);
        }

        public void SetColumnWidth(int column, double width)
        {
            ThrowIfCompleted();
            _sheet.SetColumnWidth(column, width);
        }

        /// <summary>
        /// Writes a contiguous row starting at column 1. Null entries become blank cells (skipped).
        /// </summary>
        public void WriteRow(int rowIndex, IReadOnlyList<object?> values)
        {
            ThrowIfCompleted();
            if (values is null) throw new ArgumentNullException(nameof(values));
            WriteRow(rowIndex, row =>
            {
                for (int i = 0; i < values.Count; i++)
                {
                    object? value = values[i];
                    if (value is null) continue;
                    row.Write(i + 1, value);
                }
            });
        }

        public void WriteRow(int rowIndex, Action<ExcelStreamingRowWriter> writeCells)
        {
            ThrowIfCompleted();
            if (rowIndex <= 0 || rowIndex > CellAddress.MaxRows)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), $"Row must be 1..{CellAddress.MaxRows}.");
            }

            if (writeCells is null) throw new ArgumentNullException(nameof(writeCells));

            _rowWriter.Reset(rowIndex);
            _sheet.BeginRow(rowIndex);
            writeCells(_rowWriter);
            _sheet.EndRow();
        }

        /// <summary>Closes this sheet part. Required before adding another sheet or disposing the workbook.</summary>
        public void Complete()
        {
            if (_completed) return;
            _sheet.Complete();
            _completed = true;
        }

        private void ThrowIfCompleted()
        {
            if (_completed)
            {
                throw new InvalidOperationException($"Worksheet '{Name}' is already completed.");
            }
        }
    }

    /// <summary>Cell writer for a single streaming row.</summary>
    public sealed class ExcelStreamingRowWriter
    {
        private readonly StreamingPackageWriter _package;
        private readonly StreamingSheetWriter _sheet;
        private int _rowIndex;

        internal ExcelStreamingRowWriter(StreamingPackageWriter package, StreamingSheetWriter sheet, int rowIndex)
        {
            _package = package;
            _sheet = sheet;
            _rowIndex = rowIndex;
        }

        internal void Reset(int rowIndex)
        {
            _rowIndex = rowIndex;
        }

        public void Write(int column, object? value, ExcelCellStyle? style = null)
        {
            if (value is null) return;
            switch (value)
            {
                case string s:
                    Text(column, s, style);
                    break;
                case int i:
                    Integer(column, i, style);
                    break;
                case long l:
                    Number(column, l, style);
                    break;
                case float f:
                    Number(column, f, style);
                    break;
                case double d:
                    Number(column, d, style);
                    break;
                case decimal m:
                    Number(column, (double)m, style);
                    break;
                case bool b:
                    Boolean(column, b, style);
                    break;
                case DateTime dt:
                    Date(column, dt, style);
                    break;
                case DateTimeOffset dto:
                    Date(column, dto.DateTime, style);
                    break;
                default:
                    Text(column, Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty, style);
                    break;
            }
        }

        public void Text(int column, string value, ExcelCellStyle? style = null)
        {
            EnsureColumn(column);
            _sheet.WriteStringCell(column, value ?? string.Empty, ResolveStyle(style, dateDefault: false));
        }

        public void Integer(int column, int value, ExcelCellStyle? style = null)
        {
            EnsureColumn(column);
            _sheet.WriteNumberCell(column, value.ToString(System.Globalization.CultureInfo.InvariantCulture), ResolveStyle(style, false));
        }

        public void Number(int column, double value, ExcelCellStyle? style = null)
        {
            EnsureColumn(column);
            _sheet.WriteNumberCell(
                column,
                value.ToString("G17", System.Globalization.CultureInfo.InvariantCulture),
                ResolveStyle(style, false));
        }

        public void Boolean(int column, bool value, ExcelCellStyle? style = null)
        {
            EnsureColumn(column);
            _sheet.WriteBooleanCell(column, value, ResolveStyle(style, false));
        }

        public void Date(int column, DateTime value, ExcelCellStyle? style = null)
        {
            EnsureColumn(column);
            bool useDefault = style is null || string.IsNullOrEmpty(style.NumberFormat);
            double serial = OpenXmlWorkbookWriter.ToExcelSerialDate(value);
            _sheet.WriteNumberCell(
                column,
                serial.ToString("G17", System.Globalization.CultureInfo.InvariantCulture),
                ResolveStyle(style, useDefault));
        }

        private uint? ResolveStyle(ExcelCellStyle? style, bool dateDefault)
        {
            if ((style is null || style.IsDefault) && !dateDefault)
            {
                return null;
            }

            return _package.StyleCache.GetStyleIndex(style, dateDefault);
        }

        private static void EnsureColumn(int column)
        {
            if (column <= 0 || column > CellAddress.MaxColumns)
            {
                throw new ArgumentOutOfRangeException(nameof(column), $"Column must be 1..{CellAddress.MaxColumns}.");
            }
        }
    }
}
