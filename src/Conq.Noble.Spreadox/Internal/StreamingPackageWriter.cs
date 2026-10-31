using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Conq.Noble.Spreadox.Internal
{
    /// <summary>
    /// Owns a <see cref="SpreadsheetDocument"/> and streams worksheet bodies via <see cref="OpenXmlWriter"/>.
    /// </summary>
    internal sealed class StreamingPackageWriter : IDisposable
    {
        private readonly SpreadsheetDocument _document;
        private readonly WorkbookPart _workbookPart;
        private readonly Sheets _sheets;
        private readonly Stream? _ownedStream;
        private readonly StyleIndexCache _styleCache = new StyleIndexCache();
        private readonly SharedStringTableBuilder? _sst;
        private readonly ExcelStreamingOptions _options;
        private readonly HashSet<string> _sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private ExcelStreamingWorksheet? _openSheet;
        private StreamingSheetWriter? _openSheetWriter;
        private uint _nextSheetId = 1;
        private bool _completed;

        private StreamingPackageWriter(
            SpreadsheetDocument document,
            Stream? ownedStream,
            ExcelStreamingOptions options)
        {
            _document = document;
            _ownedStream = ownedStream;
            _options = options;
            _workbookPart = document.AddWorkbookPart();
            _workbookPart.Workbook = new Workbook();
            _sheets = _workbookPart.Workbook.AppendChild(new Sheets());
            _sst = options.UseInlineStrings ? null : new SharedStringTableBuilder();
        }

        public ExcelStreamingOptions Options => _options;

        public StyleIndexCache StyleCache => _styleCache;

        public SharedStringTableBuilder? SharedStrings => _sst;

        public bool UseInlineStrings => _options.UseInlineStrings;

        public static StreamingPackageWriter Create(
            Stream stream,
            ExcelStreamingOptions options,
            bool ownsStream = false)
        {
            SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true);
            return new StreamingPackageWriter(document, ownsStream ? stream : null, options);
        }

        public ExcelStreamingWorksheet AddWorksheet(string name)
        {
            ThrowIfCompleted();
            EnsureCurrentSheetClosed();

            string trimmed = ValidateSheetName(name);
            if (!_sheetNames.Add(trimmed))
            {
                throw new InvalidOperationException($"A worksheet named '{trimmed}' already exists.");
            }

            WorksheetPart worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            string relId = _workbookPart.GetIdOfPart(worksheetPart);

            _sheets.AppendChild(new Sheet
            {
                Id = relId,
                SheetId = _nextSheetId++,
                Name = trimmed
            });

            _openSheetWriter = new StreamingSheetWriter(this, worksheetPart, trimmed);
            _openSheet = new ExcelStreamingWorksheet(this, _openSheetWriter);
            return _openSheet;
        }

        public void Dispose()
        {
            if (_completed)
            {
                return;
            }

            try
            {
                EnsureCurrentSheetClosed();

                WorkbookStylesPart stylesPart = _workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = _styleCache.BuildStylesheet();
                stylesPart.Stylesheet.Save();

                SharedStringTableBuilder? sst = _sst;
                if (sst != null)
                {
                    SharedStringTablePart sstPart = _workbookPart.AddNewPart<SharedStringTablePart>();
                    SharedStringTable table = BuildSharedStringTable(sst);
                    sstPart.SharedStringTable = table;
                    table.Save();
                }

                Workbook workbook = _workbookPart.Workbook
                    ?? throw new InvalidOperationException("Workbook part root is missing.");
                workbook.Save();
                _document.Dispose();
            }
            finally
            {
                _completed = true;
                _ownedStream?.Dispose();
            }
        }

        private void EnsureCurrentSheetClosed()
        {
            if (_openSheet != null)
            {
                _openSheet.Complete();
                _openSheet = null;
                _openSheetWriter = null;
            }
        }

        private void ThrowIfCompleted()
        {
            if (_completed)
            {
                throw new ObjectDisposedException(nameof(ExcelStreamingWorkbook));
            }
        }

        private static string ValidateSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("Worksheet name is required.", nameof(name));
            }

            string trimmed = name.Trim();
            if (trimmed.Length > 31)
            {
                throw new ArgumentException("Worksheet name cannot exceed 31 characters.", nameof(name));
            }

            if (trimmed.IndexOfAny(new[] { '\\', '/', '?', '*', '[', ']', ':' }) >= 0)
            {
                throw new ArgumentException("Worksheet name contains invalid characters.", nameof(name));
            }

            if (trimmed.StartsWith("'", StringComparison.Ordinal) || trimmed.EndsWith("'", StringComparison.Ordinal))
            {
                throw new ArgumentException("Worksheet name cannot start or end with an apostrophe.", nameof(name));
            }

            return trimmed;
        }

        private static SharedStringTable BuildSharedStringTable(SharedStringTableBuilder sst)
        {
            var table = new SharedStringTable
            {
                Count = (uint)sst.Values.Count,
                UniqueCount = (uint)sst.Values.Count
            };

            foreach (string value in sst.Values)
            {
                table.AppendChild(new SharedStringItem(new Text(OpenXmlWorkbookWriter.SanitizeXmlString(value))));
            }

            return table;
        }
    }

    /// <summary>
    /// Streams one worksheet's XML (columns + sheetData) with <see cref="OpenXmlWriter"/>.
    /// </summary>
    internal sealed class StreamingSheetWriter
    {
        private readonly StreamingPackageWriter _package;
        private readonly WorksheetPart _part;
        private readonly Dictionary<int, double> _columnWidths = new Dictionary<int, double>();
        private OpenXmlWriter? _writer;
        private bool _headerWritten;
        private bool _completed;
        private int _lastRowIndex;
        private bool _rowOpen;

        public StreamingSheetWriter(StreamingPackageWriter package, WorksheetPart part, string name)
        {
            _package = package;
            _part = part;
            Name = name;
        }

        public string Name { get; }

        public void SetColumnWidth(int column, double width)
        {
            if (_headerWritten)
            {
                throw new InvalidOperationException("Column widths must be set before writing rows.");
            }

            if (column <= 0 || column > CellAddress.MaxColumns)
            {
                throw new ArgumentOutOfRangeException(nameof(column));
            }

            if (width <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(width));
            }

            _columnWidths[column] = width;
        }

        public void BeginRow(int rowIndex)
        {
            EnsureWriter();
            if (_rowOpen)
            {
                throw new InvalidOperationException("A row is already open.");
            }

            if (rowIndex <= _lastRowIndex)
            {
                throw new InvalidOperationException(
                    $"Streaming rows must be written in ascending order (last={_lastRowIndex}, requested={rowIndex}).");
            }

            _writer!.WriteStartElement(new Row { RowIndex = (uint)rowIndex });
            _rowOpen = true;
            _lastRowIndex = rowIndex;
        }

        public void EndRow()
        {
            if (!_rowOpen)
            {
                throw new InvalidOperationException("No row is open.");
            }

            _writer!.WriteEndElement(); // Row
            _rowOpen = false;
        }

        public void WriteStringCell(int column, string value, uint? styleIndex)
        {
            EnsureRowOpen();
            string reference = CellAddress.Format(_lastRowIndex, column);
            if (_package.UseInlineStrings)
            {
                var cell = new Cell { CellReference = reference, DataType = CellValues.InlineString };
                ApplyStyle(cell, styleIndex);
                _writer!.WriteStartElement(cell);
                _writer.WriteStartElement(new InlineString());
                _writer.WriteElement(new Text(OpenXmlWorkbookWriter.SanitizeXmlString(value)));
                _writer.WriteEndElement(); // InlineString
                _writer.WriteEndElement(); // Cell
            }
            else
            {
                int index = _package.SharedStrings!.GetOrAdd(value);
                var cell = new Cell
                {
                    CellReference = reference,
                    DataType = CellValues.SharedString,
                    CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture))
                };
                ApplyStyle(cell, styleIndex);
                _writer!.WriteElement(cell);
            }
        }

        public void WriteNumberCell(int column, string invariantNumber, uint? styleIndex)
        {
            EnsureRowOpen();
            var cell = new Cell
            {
                CellReference = CellAddress.Format(_lastRowIndex, column),
                CellValue = new CellValue(invariantNumber)
            };
            ApplyStyle(cell, styleIndex);
            _writer!.WriteElement(cell);
        }

        public void WriteBooleanCell(int column, bool value, uint? styleIndex)
        {
            EnsureRowOpen();
            var cell = new Cell
            {
                CellReference = CellAddress.Format(_lastRowIndex, column),
                DataType = CellValues.Boolean,
                CellValue = new CellValue(value ? "1" : "0")
            };
            ApplyStyle(cell, styleIndex);
            _writer!.WriteElement(cell);
        }

        public void Complete()
        {
            if (_completed)
            {
                return;
            }

            if (_rowOpen)
            {
                EndRow();
            }

            EnsureWriter();
            _writer!.WriteEndElement(); // SheetData
            _writer.WriteEndElement(); // Worksheet
            _writer.Dispose();
            _writer = null;
            _completed = true;
        }

        private void EnsureWriter()
        {
            if (_completed)
            {
                throw new InvalidOperationException($"Worksheet '{Name}' is already completed.");
            }

            if (_writer != null)
            {
                return;
            }

            _writer = OpenXmlWriter.Create(_part);
            _writer.WriteStartElement(new Worksheet());
            WriteColumns();
            _writer.WriteStartElement(new SheetData());
            _headerWritten = true;
        }

        private void WriteColumns()
        {
            if (_columnWidths.Count == 0)
            {
                return;
            }

            _writer!.WriteStartElement(new Columns());
            List<int> keys = new List<int>(_columnWidths.Keys);
            keys.Sort();
            foreach (int column in keys)
            {
                _writer.WriteElement(new Column
                {
                    Min = (uint)column,
                    Max = (uint)column,
                    Width = _columnWidths[column],
                    CustomWidth = true
                });
            }

            _writer.WriteEndElement(); // Columns
        }

        private void EnsureRowOpen()
        {
            if (!_rowOpen)
            {
                throw new InvalidOperationException("Begin a row before writing cells.");
            }
        }

        private static void ApplyStyle(Cell cell, uint? styleIndex)
        {
            if (styleIndex.HasValue && styleIndex.Value != 0)
            {
                cell.StyleIndex = styleIndex.Value;
            }
        }
    }
}
