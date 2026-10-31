using System;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Public cell model. Holds typed values, optional formula, and style without exposing Open XML types.
    /// </summary>
    public sealed class ExcelCell
    {
        private object? _value;
        private CellValueKind _kind = CellValueKind.Empty;
        private string? _formula;
        private ExcelCellStyle? _style;

        public ExcelCell(ExcelWorksheet worksheet, int row, int column)
        {
            Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column));
            Row = row;
            Column = column;
        }

        public ExcelWorksheet Worksheet { get; }

        public int Row { get; }

        public int Column { get; }

        public string Address => CellAddress.Format(Row, Column);

        public CellValueKind Kind => _kind;

        /// <summary>
        /// True when there is no value and no formula (style-only cells remain "empty").
        /// </summary>
        public bool IsEmpty => !HasFormula && _kind == CellValueKind.Empty;

        public bool HasFormula => !string.IsNullOrEmpty(_formula);

        public bool HasStyle => _style != null && !_style.IsDefault;

        /// <summary>
        /// Formula text without a leading '=' (stored that way; writer adds '=' as needed).
        /// </summary>
        public string? Formula
        {
            get => _formula;
            set => SetFormula(value);
        }

        /// <summary>
        /// Lazy-created style for this cell.
        /// </summary>
        public ExcelCellStyle Style => _style ??= new ExcelCellStyle();

        public object? Value
        {
            get => _value;
            set => SetValue(value);
        }

        public void Clear()
        {
            _value = null;
            _kind = CellValueKind.Empty;
            _formula = null;
        }

        public void ClearStyle()
        {
            _style = null;
        }

        public void SetFormula(string? formula, object? cachedValue = null)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                _formula = null;
                if (cachedValue is null)
                {
                    if (_kind == CellValueKind.Formula)
                    {
                        _kind = CellValueKind.Empty;
                        _value = null;
                    }

                    return;
                }

                SetValue(cachedValue);
                return;
            }

            formula = formula!.Trim();
            if (formula.StartsWith("=", StringComparison.Ordinal))
            {
                formula = formula.Substring(1);
            }

            _formula = formula;
            _kind = CellValueKind.Formula;
            if (cachedValue != null)
            {
                ApplyCachedValueOnly(cachedValue);
            }
            else
            {
                _value = null;
            }
        }

        public void SetValue(string? value)
        {
            _formula = null;
            if (value is null)
            {
                _value = null;
                _kind = CellValueKind.Empty;
                return;
            }

            _value = value;
            _kind = CellValueKind.String;
        }

        public void SetValue(int value)
        {
            _formula = null;
            _value = value;
            _kind = CellValueKind.Integer;
        }

        public void SetValue(double value)
        {
            _formula = null;
            _value = value;
            _kind = CellValueKind.Double;
        }

        public void SetValue(decimal value)
        {
            _formula = null;
            _value = value;
            _kind = CellValueKind.Decimal;
        }

        public void SetValue(bool value)
        {
            _formula = null;
            _value = value;
            _kind = CellValueKind.Boolean;
        }

        public void SetValue(DateTime value)
        {
            _formula = null;
            _value = value;
            _kind = CellValueKind.DateTime;
        }

        public void SetValue(object? value)
        {
            if (value is null)
            {
                Clear();
                return;
            }

            switch (value)
            {
                case string s:
                    SetValue(s);
                    break;
                case int i:
                    SetValue(i);
                    break;
                case long l when l >= int.MinValue && l <= int.MaxValue:
                    SetValue((int)l);
                    break;
                case double d:
                    SetValue(d);
                    break;
                case float f:
                    SetValue((double)f);
                    break;
                case decimal m:
                    SetValue(m);
                    break;
                case bool b:
                    SetValue(b);
                    break;
                case DateTime dt:
                    SetValue(dt);
                    break;
                case DateTimeOffset dto:
                    SetValue(dto.DateTime);
                    break;
                default:
                    throw new ArgumentException(
                        $"Unsupported cell value type '{value.GetType().FullName}'. " +
                        "Supported: string, int, double, decimal, bool, DateTime, null.",
                        nameof(value));
            }
        }

        public string? GetString() =>
            _kind == CellValueKind.String || (_kind == CellValueKind.Formula && _value is string)
                ? _value as string
                : null;

        public int? GetInt() =>
            _kind == CellValueKind.Integer ? (int?)_value
            : _kind == CellValueKind.Double && _value is double d && d == Math.Truncate(d) && d >= int.MinValue && d <= int.MaxValue ? (int)d
            : _kind == CellValueKind.Formula && _value is int i ? i
            : _kind == CellValueKind.Formula && _value is double fd && fd == Math.Truncate(fd) && fd >= int.MinValue && fd <= int.MaxValue ? (int)fd
            : null;

        public double? GetDouble() =>
            _kind == CellValueKind.Double ? (double?)_value
            : _kind == CellValueKind.Integer && _value is int i ? (double)i
            : _kind == CellValueKind.Decimal && _value is decimal m ? (double)m
            : _kind == CellValueKind.Formula && _value is double d ? d
            : _kind == CellValueKind.Formula && _value is int fi ? (double)fi
            : _kind == CellValueKind.Formula && _value is decimal fm ? (double)fm
            : null;

        public decimal? GetDecimal() =>
            _kind == CellValueKind.Decimal ? (decimal?)_value
            : _kind == CellValueKind.Double && _value is double d ? (decimal)d
            : _kind == CellValueKind.Integer && _value is int i ? (decimal)i
            : _kind == CellValueKind.Formula && _value is decimal m ? m
            : _kind == CellValueKind.Formula && _value is double fd ? (decimal)fd
            : _kind == CellValueKind.Formula && _value is int fi ? (decimal)fi
            : null;

        public bool? GetBoolean() =>
            _kind == CellValueKind.Boolean ? (bool?)_value
            : _kind == CellValueKind.Formula && _value is bool b ? b
            : null;

        public DateTime? GetDateTime() =>
            _kind == CellValueKind.DateTime ? (DateTime?)_value
            : _kind == CellValueKind.Formula && _value is DateTime dt ? dt
            : null;

        internal ExcelCellStyle? StyleOrNull => _style;

        internal bool NeedsPersist => HasFormula || HasStyle || _kind != CellValueKind.Empty;

        internal void Load(CellValueKind kind, object? value, string? formula, ExcelCellStyle? style)
        {
            _kind = kind;
            _value = value;
            _formula = formula;
            _style = style;
        }

        private void ApplyCachedValueOnly(object value)
        {
            switch (value)
            {
                case string s:
                    _value = s;
                    break;
                case int i:
                    _value = i;
                    break;
                case double d:
                    _value = d;
                    break;
                case float f:
                    _value = (double)f;
                    break;
                case decimal m:
                    _value = m;
                    break;
                case bool b:
                    _value = b;
                    break;
                case DateTime dt:
                    _value = dt;
                    break;
                default:
                    _value = value.ToString();
                    break;
            }
        }
    }
}
