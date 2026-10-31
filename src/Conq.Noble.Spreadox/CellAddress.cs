using System;
using System.Globalization;
using System.Text;

namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Parses and formats Excel A1-style addresses (e.g. A1, AB12). Rows and columns are 1-based.
    /// </summary>
    public static class CellAddress
    {
        /// <summary>Excel worksheet row limit (1-based inclusive).</summary>
        public const int MaxRows = 1_048_576;

        /// <summary>Excel worksheet column limit (XFD = 16384).</summary>
        public const int MaxColumns = 16_384;

        public static void Parse(string a1, out int row, out int column)
        {
            if (string.IsNullOrWhiteSpace(a1))
            {
                throw new ArgumentException("Cell address is required.", nameof(a1));
            }

            a1 = a1.Trim().ToUpperInvariant();
            int i = 0;
            int col = 0;
            while (i < a1.Length && a1[i] >= 'A' && a1[i] <= 'Z')
            {
                col = (col * 26) + (a1[i] - 'A' + 1);
                i++;
            }

            if (col <= 0 || i >= a1.Length)
            {
                throw new ArgumentException($"Invalid cell address '{a1}'.", nameof(a1));
            }

            if (!int.TryParse(a1.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out int r) || r <= 0)
            {
                throw new ArgumentException($"Invalid cell address '{a1}'.", nameof(a1));
            }

            if (r > MaxRows || col > MaxColumns)
            {
                throw new ArgumentException($"Cell address '{a1}' exceeds sheet bounds (MaxRows={MaxRows}, MaxColumns={MaxColumns}).", nameof(a1));
            }

            row = r;
            column = col;
        }

        public static string Format(int row, int column)
        {
            if (row <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(row));
            }

            if (column <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(column));
            }

            var letters = new StringBuilder();
            int c = column;
            while (c > 0)
            {
                c--;
                letters.Insert(0, (char)('A' + (c % 26)));
                c /= 26;
            }

            return letters.ToString() + row.ToString(CultureInfo.InvariantCulture);
        }
    }
}
