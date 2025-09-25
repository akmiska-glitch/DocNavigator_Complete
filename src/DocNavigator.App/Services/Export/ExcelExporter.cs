using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using ClosedXML.Excel;

namespace DocNavigator.App.Services.Export
{
    public sealed class ExportProgress
    {
        public int SheetIndex { get; set; }        // 1-based
        public int SheetCount { get; set; }
        public string SheetName { get; set; } = string.Empty;
        public int RowsDone { get; set; }
        public int RowsTotal { get; set; }
        public string Phase { get; set; } = string.Empty; // "Подготовка" / "Запись" / "Автоширина"
    }

    public static class ExcelExporter
    {
        /// <summary>
        /// Экспорт одной таблицы. Автоширина: шапка + первые N строк (по умолчанию 200).
        /// </summary>
        public static void SaveSingleTable(DataTable table, string sheetName, string filePath, int autoFitSampleRows = 200, IProgress<ExportProgress>? progress = null, CancellationToken ct = default)
        {
            using (var wb = new XLWorkbook())
            {
                AddWorksheet(wb, table, sheetName, autoFitSampleRows, progress, 1, 1, ct);
                SaveWorkbook(wb, filePath);
            }
        }

        /// <summary>
        /// Экспорт нескольких таблиц. Каждый элемент: (SheetName, DataTable).
        /// </summary>
        public static void SaveMultipleTables((string Sheet, DataTable Table)[] sheets, string filePath, int autoFitSampleRows = 200, IProgress<ExportProgress>? progress = null, CancellationToken ct = default)
        {
            using (var wb = new XLWorkbook())
            {
                int idx = 0;
                foreach (var s in sheets)
                {
                    ct.ThrowIfCancellationRequested();
                    idx++;
                    AddWorksheet(wb, s.Table, s.Sheet, autoFitSampleRows, progress, idx, sheets.Length, ct);
                }
                SaveWorkbook(wb, filePath);
            }
        }

        private static void AddWorksheet(XLWorkbook wb, DataTable table, string sheetName, int autoFitSampleRows, IProgress<ExportProgress>? progress, int sheetIndex, int sheetCount, CancellationToken ct)
        {
            var safeName = MakeSafeSheetName(sheetName);
            if (string.IsNullOrWhiteSpace(safeName)) safeName = $"Sheet{sheetIndex}";
            if (wb.Worksheets.Any(ws => string.Equals(ws.Name, safeName, StringComparison.OrdinalIgnoreCase)))
            {
                int i = 2;
                var baseName = safeName.Length > 28 ? safeName.Substring(0, 28) : safeName;
                while (wb.Worksheets.Any(ws => string.Equals(ws.Name, safeName, StringComparison.OrdinalIgnoreCase)))
                    safeName = $"{baseName}_{i++}";
            }

            progress?.Report(new ExportProgress { SheetIndex = sheetIndex, SheetCount = sheetCount, SheetName = safeName, RowsDone = 0, RowsTotal = table.Rows.Count, Phase = "Подготовка" });

            var ws = wb.Worksheets.Add(safeName);

            ct.ThrowIfCancellationRequested();
            progress?.Report(new ExportProgress { SheetIndex = sheetIndex, SheetCount = sheetCount, SheetName = safeName, RowsTotal = table.Rows.Count, RowsDone = 0, Phase = "Запись" });

            // Вставляем DataTable одним вызовом (быстро)
            var range = ws.Cell(1, 1).InsertTable(table, true).AsRange();

            // Автофильтр в шапке (если метод недоступен в вашей версии ClosedXML — игнорируем исключение)
            try { range.SetAutoFilter(); } catch { }

            // Форматы по типам
            ApplyColumnDataTypes(range, table);

            // Автоширина по шапке + первым N строкам (ускоряет большие наборы)
            ct.ThrowIfCancellationRequested();
            progress?.Report(new ExportProgress { SheetIndex = sheetIndex, SheetCount = sheetCount, SheetName = safeName, RowsTotal = table.Rows.Count, RowsDone = table.Rows.Count, Phase = "Автоширина" });
            AdjustColumnsSample(range, autoFitSampleRows);

            // Заморозить шапку
            try { ws.SheetView.FreezeRows(1); } catch { }
        }

        private static void ApplyColumnDataTypes(IXLRange range, DataTable table)
        {
            for (int i = 0; i < table.Columns.Count; i++)
            {
                var dc = table.Columns[i];
                var col = range.Worksheet.Column(range.FirstColumn().ColumnNumber() + i);

                if (dc.DataType == typeof(DateTime) || dc.DataType == typeof(DateTimeOffset))
                {
                    col.Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
                }
                else if (dc.DataType == typeof(decimal) || dc.DataType == typeof(double) || dc.DataType == typeof(float))
                {
                    col.Style.NumberFormat.Format = "0.############";
                }
                else if (dc.DataType == typeof(int) || dc.DataType == typeof(long) || dc.DataType == typeof(short))
                {
                    col.Style.NumberFormat.Format = "0";
                }
                else
                {
                    col.Style.NumberFormat.Format = "@"; // текст: не терять ведущие нули и длинные идентификаторы
                }
            }
        }

        private static void AdjustColumnsSample(IXLRange range, int sampleRows)
        {
            var ws = range.Worksheet;

            int firstCol = range.FirstColumn().ColumnNumber();
            int lastCol  = range.LastColumn().ColumnNumber();

            int headerRow    = range.FirstRow().RowNumber();
            int firstDataRow = headerRow + 1;
            int lastDataRow  = Math.Min(firstDataRow + Math.Max(sampleRows, 0) - 1, range.LastRow().RowNumber());

            for (int colNum = firstCol; colNum <= lastCol; colNum++)
            {
                var col = ws.Column(colNum);
                try { col.AdjustToContents(headerRow, headerRow); } catch { }
                if (lastDataRow >= firstDataRow)
                {
                    try { col.AdjustToContents(firstDataRow, lastDataRow); } catch { }
                }
            }
        }

        private static void SaveWorkbook(XLWorkbook wb, string filePath)
        {
            try
            {
                wb.SaveAs(filePath);
            }
            catch (IOException)
            {
                var dir  = Path.GetDirectoryName(filePath) ?? "";
                var name = Path.GetFileNameWithoutExtension(filePath);
                var ext  = Path.GetExtension(filePath);
                var alt  = Path.Combine(dir, $"{name}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");
                wb.SaveAs(alt);
            }
        }

        private static string MakeSafeSheetName(string name)
        {
            var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (var ch in invalid)
                name = name.Replace(ch, '_');
            if (name.Length > 31) name = name.Substring(0, 31);
            return name.Trim();
        }
    }
}
