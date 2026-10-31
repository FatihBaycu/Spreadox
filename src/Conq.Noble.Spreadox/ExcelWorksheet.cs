using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// In-memory worksheet. Cell storage is keyed by 1-based row/column.
    /// </summary>
    public sealed class ExcelWorksheet
    {
        private readonly Dictionary<(int Row, int Column), ExcelCell> _cells =
            new Dictionary<(int, int), ExcelCell>();

        private readonly Dictionary<int, double> _rowHeights = new Dictionary<int, double>();
        private readonly Dictionary<int, double> _columnWidths = new Dictionary<int, double>();
        private readonly List<ExcelRange> _mergedRanges = new List<ExcelRange>();
        private readonly List<ExcelConditionalFormat> _conditionalFormats = new List<ExcelConditionalFormat>();

        private string _name;
        private string? _autoFilterRange;
        private int _freezeRows;
        private int _freezeColumns;

        internal ExcelWorksheet(ExcelWorkbook workbook, string name)
        {
            Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _name = ValidateName(name);
            Charts = new ExcelChartCollection(this);
            Pictures = new ExcelPictureCollection(this);
        }

        public ExcelWorkbook Workbook { get; }

        /// <summary>Native OOXML charts on this worksheet.</summary>
        public ExcelChartCollection Charts { get; }

        /// <summary>Pictures anchored to cell ranges.</summary>
        public ExcelPictureCollection Pictures { get; }

        public string Name
        {
            get => _name;
            set
            {
                Workbook.ThrowIfDisposed();
                string newName = ValidateName(value);
                if (string.Equals(_name, newName, StringComparison.OrdinalIgnoreCase))
                {
                    _name = newName;
                    return;
                }

                Workbook.EnsureWorksheetNameAvailable(newName, this);
                _name = newName;
            }
        }

        public IReadOnlyList<ExcelRange> MergedRanges => _mergedRanges;

        public IReadOnlyList<ExcelConditionalFormat> ConditionalFormats => _conditionalFormats;

        /// <summary>AutoFilter range in A1 notation (e.g. A1:D10), or null.</summary>
        public string? AutoFilterRange
        {
            get => _autoFilterRange;
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    _autoFilterRange = null;
                    return;
                }

                _autoFilterRange = ExcelRange.Parse(value).A1;
            }
        }

        public int FreezeRows => _freezeRows;

        public int FreezeColumns => _freezeColumns;

        public ExcelCell this[string a1Address]
        {
            get
            {
                Workbook.ThrowIfDisposed();
                if (a1Address is null) throw new ArgumentNullException(nameof(a1Address));
                CellAddress.Parse(a1Address, out int row, out int column);
                return this[row, column];
            }
        }

        public ExcelCell this[int row, int column]
        {
            get
            {
                Workbook.ThrowIfDisposed();
                if (row <= 0 || row > CellAddress.MaxRows)
                {
                    throw new ArgumentOutOfRangeException(nameof(row), $"Row must be 1..{CellAddress.MaxRows}.");
                }

                if (column <= 0 || column > CellAddress.MaxColumns)
                {
                    throw new ArgumentOutOfRangeException(nameof(column), $"Column must be 1..{CellAddress.MaxColumns}.");
                }

                var key = (row, column);
                if (!_cells.TryGetValue(key, out ExcelCell? cell))
                {
                    cell = new ExcelCell(this, row, column);
                    _cells[key] = cell;
                }

                return cell;
            }
        }

        public IReadOnlyCollection<ExcelCell> Cells =>
            new List<ExcelCell>(_cells.Values);

        public void SetValue(string a1Address, object? value) => this[a1Address].SetValue(value);

        public object? GetValue(string a1Address)
        {
            CellAddress.Parse(a1Address, out int row, out int column);
            return TryGetCell(row, column, out ExcelCell? cell) ? cell!.Value : null;
        }

        public void SetFormula(string a1Address, string formula, object? cachedValue = null) =>
            this[a1Address].SetFormula(formula, cachedValue);

        public void SetRowHeight(int row, double heightPoints)
        {
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (heightPoints <= 0) throw new ArgumentOutOfRangeException(nameof(heightPoints));
            _rowHeights[row] = heightPoints;
        }

        public bool TryGetRowHeight(int row, out double heightPoints) =>
            _rowHeights.TryGetValue(row, out heightPoints);

        public void SetColumnWidth(int column, double width)
        {
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column));
            if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
            _columnWidths[column] = width;
        }

        public void SetColumnWidth(string columnLetters, double width)
        {
            int column = ParseColumnLetters(columnLetters);
            SetColumnWidth(column, width);
        }

        public bool TryGetColumnWidth(int column, out double width) =>
            _columnWidths.TryGetValue(column, out width);

        public void Merge(string a1Range) => Merge(ExcelRange.Parse(a1Range));

        public void Merge(ExcelRange range)
        {
            foreach (ExcelRange existing in _mergedRanges)
            {
                if (RangesOverlap(existing, range))
                {
                    throw new InvalidOperationException(
                        $"Merge range {range.A1} overlaps existing merge {existing.A1}.");
                }
            }

            _mergedRanges.Add(range);
        }

        public bool Unmerge(string a1Range)
        {
            ExcelRange range = ExcelRange.Parse(a1Range);
            for (int i = 0; i < _mergedRanges.Count; i++)
            {
                if (_mergedRanges[i].Equals(range))
                {
                    _mergedRanges.RemoveAt(i);
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Freezes the first <paramref name="rows"/> rows and <paramref name="columns"/> columns.
        /// </summary>
        public void FreezePanes(int rows, int columns = 0)
        {
            if (rows < 0) throw new ArgumentOutOfRangeException(nameof(rows));
            if (columns < 0) throw new ArgumentOutOfRangeException(nameof(columns));
            _freezeRows = rows;
            _freezeColumns = columns;
        }

        public void UnfreezePanes()
        {
            _freezeRows = 0;
            _freezeColumns = 0;
        }

        public void AddConditionalFormat(ExcelConditionalFormat rule)
        {
            if (rule is null) throw new ArgumentNullException(nameof(rule));
            _conditionalFormats.Add(rule);
        }

        public void AddConditionalFormat(
            string rangeA1,
            ConditionalOperator op,
            string formulaOperand,
            ExcelColor? fillColor = null)
        {
            AddConditionalFormat(new ExcelConditionalFormat(rangeA1, op, formulaOperand, fillColor));
        }

        /// <summary>3-color percentile color scale (0 / 50 / 100).</summary>
        public ExcelConditionalFormat AddColorScale(
            string rangeA1,
            ExcelColor low,
            ExcelColor midpoint,
            ExcelColor high)
        {
            ExcelConditionalFormat rule = ExcelConditionalFormat.ColorScalePercentile(rangeA1, low, midpoint, high);
            AddConditionalFormat(rule);
            return rule;
        }

        /// <summary>
        /// Inserts rows of cell values starting at the given cell (1-based).
        /// </summary>
        public void Insert(IEnumerable<object?[]> rows, int startRow = 1, int startColumn = 1)
        {
            if (rows is null) throw new ArgumentNullException(nameof(rows));
            if (startRow <= 0) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn <= 0) throw new ArgumentOutOfRangeException(nameof(startColumn));

            int r = startRow;
            foreach (object?[] row in rows)
            {
                if (row is null)
                {
                    r++;
                    continue;
                }

                for (int i = 0; i < row.Length; i++)
                {
                    this[r, startColumn + i].SetValue(row[i]);
                }

                r++;
            }
        }

        public void Insert(IEnumerable<IEnumerable> rows, int startRow = 1, int startColumn = 1)
        {
            if (rows is null) throw new ArgumentNullException(nameof(rows));
            var materialized = new List<object?[]>();
            foreach (IEnumerable row in rows)
            {
                if (row is null)
                {
                    materialized.Add(Array.Empty<object?>());
                    continue;
                }

                var values = new List<object?>();
                foreach (object? item in row)
                {
                    values.Add(item);
                }

                materialized.Add(values.ToArray());
            }

            Insert(materialized, startRow, startColumn);
        }

        public void InsertDataTable(DataTable table, int startRow = 1, int startColumn = 1, bool includeHeaders = true)
        {
            if (table is null) throw new ArgumentNullException(nameof(table));
            if (startRow <= 0) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn <= 0) throw new ArgumentOutOfRangeException(nameof(startColumn));

            int r = startRow;
            if (includeHeaders)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    this[r, startColumn + c].SetValue(table.Columns[c].ColumnName);
                }

                r++;
            }

            foreach (DataRow dataRow in table.Rows)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    object value = dataRow[c];
                    this[r, startColumn + c].SetValue(value == DBNull.Value ? null : value);
                }

                r++;
            }
        }

        /// <summary>
        /// Projects public readable properties of <typeparamref name="T"/> into columns.
        /// </summary>
        public void InsertObjects<T>(IEnumerable<T> items, int startRow = 1, int startColumn = 1, bool includeHeaders = true)
        {
            if (items is null) throw new ArgumentNullException(nameof(items));
            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);
            var readable = new List<PropertyInfo>();
            foreach (PropertyInfo prop in props)
            {
                if (prop.CanRead && prop.GetIndexParameters().Length == 0)
                {
                    readable.Add(prop);
                }
            }

            int r = startRow;
            if (includeHeaders)
            {
                for (int c = 0; c < readable.Count; c++)
                {
                    this[r, startColumn + c].SetValue(readable[c].Name);
                }

                r++;
            }

            foreach (T item in items)
            {
                if (item is null)
                {
                    r++;
                    continue;
                }

                for (int c = 0; c < readable.Count; c++)
                {
                    this[r, startColumn + c].SetValue(readable[c].GetValue(item));
                }

                r++;
            }
        }

        internal bool TryGetCell(int row, int column, out ExcelCell? cell) =>
            _cells.TryGetValue((row, column), out cell);

        internal IEnumerable<ExcelCell> EnumeratePersistableCells()
        {
            foreach (ExcelCell cell in _cells.Values)
            {
                if (cell.NeedsPersist)
                {
                    yield return cell;
                }
            }
        }

        internal IReadOnlyDictionary<int, double> RowHeights => _rowHeights;

        internal IReadOnlyDictionary<int, double> ColumnWidths => _columnWidths;

        internal void LoadMergedRange(ExcelRange range) => _mergedRanges.Add(range);

        internal void LoadRowHeight(int row, double height) => _rowHeights[row] = height;

        internal void LoadColumnWidth(int column, double width) => _columnWidths[column] = width;

        internal void LoadFreezePanes(int rows, int columns)
        {
            _freezeRows = rows;
            _freezeColumns = columns;
        }

        internal void LoadAutoFilter(string? range) => _autoFilterRange = range;

        internal void LoadConditionalFormat(ExcelConditionalFormat rule) => _conditionalFormats.Add(rule);

        private static bool RangesOverlap(ExcelRange a, ExcelRange b) =>
            a.StartRow <= b.EndRow
            && a.EndRow >= b.StartRow
            && a.StartColumn <= b.EndColumn
            && a.EndColumn >= b.StartColumn;

        private static int ParseColumnLetters(string columnLetters)
        {
            if (string.IsNullOrWhiteSpace(columnLetters))
            {
                throw new ArgumentException("Column letters are required.", nameof(columnLetters));
            }

            string text = columnLetters.Trim().ToUpperInvariant();
            int col = 0;
            foreach (char ch in text)
            {
                if (ch < 'A' || ch > 'Z')
                {
                    throw new ArgumentException($"Invalid column '{columnLetters}'.", nameof(columnLetters));
                }

                col = (col * 26) + (ch - 'A' + 1);
            }

            return col;
        }

        private static string ValidateName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("Worksheet name is required.", nameof(name));
            }

            name = name.Trim();
            if (name.Length > 31)
            {
                throw new ArgumentException("Worksheet name cannot exceed 31 characters.", nameof(name));
            }

            if (name.IndexOfAny(new[] { '\\', '/', '?', '*', '[', ']', ':' }) >= 0)
            {
                throw new ArgumentException("Worksheet name contains invalid characters.", nameof(name));
            }

            if (name.StartsWith("'", StringComparison.Ordinal) || name.EndsWith("'", StringComparison.Ordinal))
            {
                throw new ArgumentException("Worksheet name cannot start or end with an apostrophe.", nameof(name));
            }

            return name;
        }
    }
}
