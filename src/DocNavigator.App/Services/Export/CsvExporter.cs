// File: src/DocNavigator.App/Services/Export/CsvExporter.cs
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace DocNavigator.App.Services.Export
{
    public static class CsvExporter
    {
        /// <summary>
        /// Сохраняет DataTable как "сырые" данные (без локализаций/форматов) в CSV (UTF-8, разделитель — запятая).
        /// </summary>
        public static void Save(DataTable table, string filePath, char delimiter = ',')
        {
            using var sw = new StreamWriter(filePath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            // Заголовки
            for (int c = 0; c < table.Columns.Count; c++)
            {
                if (c > 0) sw.Write(delimiter);
                sw.Write(Escape(table.Columns[c].ColumnName, delimiter));
            }
            sw.WriteLine();

            // Данные
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    if (c > 0) sw.Write(delimiter);
                    var val = table.Rows[r][c];
                    sw.Write(Escape(val, delimiter));
                }
                sw.WriteLine();
            }
        }

        private static string Escape(object? value, char delimiter)
        {
            if (value == null || value == DBNull.Value) return "";
            var s = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
            var needsQuotes = s.Contains(delimiter) || s.Contains('\n') || s.Contains('\r') || s.Contains('"');
            if (needsQuotes)
            {
                s = s.Replace("\"", "\"\"");
                s = $"\"{s}\"";
            }
            return s;
        }
    }
}
