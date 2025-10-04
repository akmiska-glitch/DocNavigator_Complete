using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using ClosedXML.Excel;
using DocNavigator.App.Models;
using DocNavigator.App.Services.Metadata;
using System.Collections;
namespace DocNavigator.App.Services.Export
{
    public sealed class ExportProgress
    {
        public int SheetIndex { get; set; }        // 1-based
        public int SheetCount { get; set; }
        public string SheetName { get; set; } = string.Empty;
        public string Phase { get; set; } = string.Empty; // Подготовка / Запись / Автоширина
    }

    public static class ExcelExporter
    {
        /// <summary>Сохранить одну таблицу в XLSX.</summary>
        public static void SaveSingleTable(
            DataTable table,
            string sheetName,
            string filePath,
            int autoFitSampleRows = 200,
            IProgress<ExportProgress>? progress = null,
            CancellationToken ct = default,
            DescriptorMeta? meta = null)
        {
            using (var wb = new XLWorkbook())
            {
                AddWorksheet(wb, table, sheetName, autoFitSampleRows, progress, 1, 1, ct, meta);
                SaveWorkbook(wb, filePath);
            }
        }

        /// <summary>Сохранить несколько таблиц в один XLSX (каждая в своём листе).</summary>
        public static void SaveMultipleTables(
            (string Sheet, DataTable Table)[] sheets,
            string filePath,
            int autoFitSampleRows = 200,
            IProgress<ExportProgress>? progress = null,
            CancellationToken ct = default,
            DescriptorMeta? meta = null)
        {
            using (var wb = new XLWorkbook())
            {
                for (int i = 0; i < sheets.Length; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    var s = sheets[i];
                    AddWorksheet(wb, s.Table, s.Sheet, autoFitSampleRows, progress, i + 1, sheets.Length, ct, meta);
                }
                SaveWorkbook(wb, filePath);
            }
        }

        private static void AddWorksheet(
            XLWorkbook wb,
            DataTable table,
            string sheetName,
            int autoFitSampleRows,
            IProgress<ExportProgress>? progress,
            int sheetIndex,
            int sheetCount,
            CancellationToken ct,
            DescriptorMeta? meta)
        {
            progress?.Report(new ExportProgress { SheetIndex = sheetIndex, SheetCount = sheetCount, SheetName = sheetName, Phase = "Подготовка" });

            var ws = wb.Worksheets.Add(MakeSafeSheetName(sheetName));

            // 1) План типизации по .desc и правилам
            var plan = BuildTypingPlan(table, meta, out var debugRows);

            // 2) Формируем таблицу с нужными CLR-типами (string / decimal / DateTime)
            var exportTable = MaterializeForExcel(table, plan);

            // 3) Вставляем одним вызовом
            var range = ws.Cell(1, 1).InsertTable(exportTable, true).AsRange();
            range.Style.NumberFormat.Format = "General";
            try { range.SetAutoFilter(); } catch { }

            // 4) Оформление шапки
            var header = range.FirstRow();
            header.Style.Font.Bold = true;
            header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // 5) Форматы/фиксация текстов/дат/чисел
            ApplyColumnStyles(range, plan);

            // 6) Автоширина по «хвосту» + высота шапки
            progress?.Report(new ExportProgress { SheetIndex = sheetIndex, SheetCount = sheetCount, SheetName = sheetName, Phase = "Автоширина" });
            AdjustColumnsSample(range, autoFitSampleRows);
            try { ws.Row(header.RowNumber()).AdjustToContents(); } catch { }

            // 7) Видимый отладочный лист
            try { AppendDebugSheetVisible(wb, sheetName, plan, debugRows); } catch { }
        }

        // ===== Типизация: по .desc (+ даты в doc/routecontext), без «угадывания чисел» =====

        private enum TargetKind { Text, Decimal, DateTime }

        private sealed class ColumnPlan
        {
            public int Index;                // 0-based
            public string ColumnName = "";
            public string SysName = "";
            public string DescType = "";     // из .desc (нормализованный)
            public TargetKind Kind;
            public bool ForcedTextByStringDesc;
            public bool ForcedTextByTableRule;   // doc/routecontext non-date, DC/FS id-колонки
            public bool HeuristicDate;           // только для doc/routecontext
        }

        private static List<ColumnPlan> BuildTypingPlan(DataTable source, DescriptorMeta? meta, out List<string[]> debugRows)
        {
            string tableName = (source.TableName ?? "").ToLowerInvariant();
            bool isDocOrRoute = tableName == "doc" || tableName == "routecontext";
            bool isDcOrFs = tableName.StartsWith("dc_") || tableName.StartsWith("fs_");

            bool IsAlwaysTextId(string sys)
            {
                switch ((sys ?? "").ToLowerInvariant())
                {
                    case "docid":
                    case "version":
                    case "tablerownum":
                    case "fieldsetid":
                        return true;
                    default:
                        return false;
                }
            }

            var plans = new List<ColumnPlan>(source.Columns.Count);
            debugRows = new List<string[]>();

            for (int i = 0; i < source.Columns.Count; i++)
            {
                var dc = source.Columns[i];

                // sys-имя
                string sys = (dc.ExtendedProperties["sys"] as string)
                             ?? (dc.ExtendedProperties["systemName"] as string)
                             ?? (dc.ExtendedProperties["fieldSystemName"] as string)
                             ?? dc.ColumnName;

                // .desc: надёжное разрешение типа
string kindDesc = (ResolveDescType(meta, sys) ?? "").Trim().ToUpperInvariant();



                // Эвристика даты — только для doc/routecontext
                bool looksDate = isDocOrRoute && LooksLikeDate(source, dc);

                var plan = new ColumnPlan
                {
                    Index = i,
                    ColumnName = dc.ColumnName,
                    SysName = sys,
                    DescType = kindDesc,
                    ForcedTextByStringDesc = kindDesc == "STRING",
                    ForcedTextByTableRule = (isDocOrRoute && !(kindDesc == "DATE" || dc.DataType == typeof(DateTime))) // всё, что не дата
                                             || (isDcOrFs && IsAlwaysTextId(sys)),
                    HeuristicDate = looksDate
                };

                // Итоговый тип:
                if (plan.ForcedTextByStringDesc || plan.ForcedTextByTableRule)
                {
                    plan.Kind = TargetKind.Text;
                }
                else if (kindDesc == "DATE" || dc.DataType == typeof(DateTime) || plan.HeuristicDate)
                {
                    plan.Kind = TargetKind.DateTime;
                }
                else if (kindDesc == "DECIMAL")
                {
                    plan.Kind = TargetKind.Decimal;
                }
                else
                {
                    plan.Kind = TargetKind.Text;
                }

                plans.Add(plan);

                // строка дебага
                debugRows.Add(new[]
                {
                    source.TableName ?? "",
                    dc.ColumnName,
                    sys,
                    kindDesc,
                    plan.Kind.ToString(),
                    plan.ForcedTextByStringDesc ? "Y" : "",
                    plan.ForcedTextByTableRule ? "Y" : "",
                    plan.HeuristicDate ? "Y" : ""
                });
            }

            return plans;
        }

        /// <summary>
        /// Унифицированное разрешение типа поля по .desc (устойчиво к регистру/префиксам).
        /// </summary>
        private static string? ResolveDescType(DescriptorMeta? meta, string sys)
        {
            if (meta?.FieldsBySystemName == null) return null;
            var dict = meta.FieldsBySystemName;

            if (string.IsNullOrWhiteSpace(sys)) return null;
            var key0 = sys.Trim();
            var keyU = key0.ToUpperInvariant();
            var keyL = key0.ToLowerInvariant();

            // 1) прямые попытки
            if (dict.TryGetValue(key0, out var fm1) && fm1?.DataType != null) return fm1.DataType;
            if (dict.TryGetValue(keyU, out var fm2) && fm2?.DataType != null) return fm2.DataType;
            if (dict.TryGetValue(keyL, out var fm3) && fm3?.DataType != null) return fm3.DataType;

            // 2) кейс-инсенситивный перебор ключей
            foreach (var kv in dict)
                if (string.Equals(kv.Key?.Trim(), key0, StringComparison.OrdinalIgnoreCase))
                    return kv.Value?.DataType;

            // 3) если ключи содержат префикс (например, "FS_S1_T1_0531444_LIST:C4" или ".../C4")
            foreach (var kv in dict)
            {
                var k = kv.Key?.Trim() ?? "";
                if (k.EndsWith(":" + key0, StringComparison.OrdinalIgnoreCase) ||
                    k.EndsWith("/" + key0, StringComparison.OrdinalIgnoreCase))
                    return kv.Value?.DataType;
            }

            // 4) по значению SystemName в FieldMeta (если доступно)
            foreach (var kv in dict)
            {
                var fm = kv.Value;
                // пытаемся прочитать SystemName отражением на случай другой модели
                string? sysName = null;
                try
                {
                    var prop = fm?.GetType().GetProperty("SystemName");
                    if (prop != null)
                        sysName = prop.GetValue(fm)?.ToString();
                }
                catch { /* ignore */ }

                if (!string.IsNullOrEmpty(sysName) &&
                    string.Equals(sysName.Trim(), key0, StringComparison.OrdinalIgnoreCase))
                    return fm?.DataType;
            }

            return null;
        }

        private static DataTable MaterializeForExcel(DataTable source, List<ColumnPlan> plan)
        {
            var dst = new DataTable(source.TableName);

            // Создаём колонки с нужными CLR-типами
            foreach (var p in plan)
            {
                Type t = p.Kind == TargetKind.Text ? typeof(string)
                       : p.Kind == TargetKind.Decimal ? typeof(decimal)
                       : typeof(DateTime);

                var srcCol = source.Columns[p.Index];
                var newCol = new DataColumn(srcCol.ColumnName, t);
                foreach (var key in srcCol.ExtendedProperties.Keys)
                    newCol.ExtendedProperties[key] = srcCol.ExtendedProperties[key];
                dst.Columns.Add(newCol);
            }

            // Перенос значений
            foreach (DataRow src in source.Rows)
            {
                var row = dst.NewRow();
                for (int i = 0; i < plan.Count; i++)
                {
                    var p = plan[i];
                    var srcCol = source.Columns[p.Index];
                    object val = src[srcCol];

                    if (val == DBNull.Value)
                    {
                        row[i] = DBNull.Value;
                        continue;
                    }

                    switch (p.Kind)
                    {
                        case TargetKind.Text:
                            row[i] = Convert.ToString(val, CultureInfo.InvariantCulture) ?? string.Empty;
                            break;

                        case TargetKind.DateTime:
                            if (val is DateTime dt)
                            {
                                row[i] = dt;
                            }
                            else
                            {
                                var s = Convert.ToString(val, CultureInfo.InvariantCulture);
                                if (!string.IsNullOrWhiteSpace(s) && TryParseDateFlexible(s, out var parsed))
                                    row[i] = parsed;
                                else
                                    row[i] = DBNull.Value;
                            }
                            break;

                        case TargetKind.Decimal:
                            if (val is decimal de) row[i] = de;
                            else if (val is double dd) row[i] = Convert.ToDecimal(dd);
                            else if (val is float ff) row[i] = Convert.ToDecimal(ff);
                            else
                            {
                                var s = Convert.ToString(val, CultureInfo.InvariantCulture);
                                if (!string.IsNullOrWhiteSpace(s))
                                {
                                    s = s.Replace(" ", "").Replace("\u00A0", "");
                                    if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var pdec)
                                     || decimal.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out pdec))
                                        row[i] = pdec;
                                    else
                                        row[i] = DBNull.Value;
                                }
                                else row[i] = DBNull.Value;
                            }
                            break;
                    }
                }
                dst.Rows.Add(row);
            }

            return dst;
        }

        private static void ApplyColumnStyles(IXLRange range, List<ColumnPlan> plan)
        {
            var ws = range.Worksheet;
            int headerRow = range.FirstRow().RowNumber();
            int firstCol = range.FirstColumn().ColumnNumber();
            int lastCol = range.LastColumn().ColumnNumber();
            int firstDataRow = headerRow + 1;
            int lastRow = range.LastRow().RowNumber();

            for (int i = 0; i < plan.Count; i++)
            {
                var p = plan[i];
                int colNum = firstCol + i;
                var col = ws.Column(colNum);

                if (p.Kind == TargetKind.DateTime)
                {
                    col.Style.DateFormat.Format = "dd.MM.yyyy HH:mm:ss";
                }
                else if (p.Kind == TargetKind.Decimal)
                {
                    col.Style.NumberFormat.Format = "#,##0.00";
                }
                else // Text
                {
                    col.Style.NumberFormat.Format = "@";
                    // добавим апостроф для всей колонки (кроме пустых), чтобы Excel не переводил в E+ и не отбрасывал нули
                    if (firstDataRow <= lastRow)
                    {
                        for (int r = firstDataRow; r <= lastRow; r++)
                        {
                            var cell = ws.Cell(r, colNum);
                            if (cell.IsEmpty()) continue;

                            var s = cell.GetString();
                            if (!string.IsNullOrEmpty(s) && s[0] != '\'')
                                cell.Value = "'" + s;
                        }
                    }
                }
            }
        }

        // ===== Автоширина =====

        private static void AdjustColumnsSample(IXLRange range, int sampleRows)
        {
            var ws = range.Worksheet;

            int firstCol = range.FirstColumn().ColumnNumber();
            int lastCol = range.LastColumn().ColumnNumber();

            int headerRow = range.FirstRow().RowNumber();
            int firstDataRow = headerRow + 1;
            int lastRow = range.LastRow().RowNumber();

            if (firstDataRow > lastRow)
            {
                for (int colNum = firstCol; colNum <= lastCol; colNum++)
                {
                    var c = ws.Column(colNum);
                    try { c.AdjustToContents(headerRow, headerRow); } catch { }
                    c.Width = Math.Min(255.0, Math.Max(c.Width, 8.0) + 1.5);
                    c.Style.Alignment.ShrinkToFit = false;
                }
                return;
            }

            int sampleStart = Math.Max(firstDataRow, lastRow - Math.Max(sampleRows, 0) + 1);
            int sampleEnd = lastRow;

            for (int colNum = firstCol; colNum <= lastCol; colNum++)
            {
                var c = ws.Column(colNum);

                double widthHeader, widthData;

                try { c.AdjustToContents(headerRow, headerRow); } catch { }
                widthHeader = c.Width;

                try { c.AdjustToContents(sampleStart, sampleEnd); } catch { }
                widthData = c.Width;

                double target = Math.Max(widthData, widthHeader);
                if (target < 8.0) target = 8.0;
                target = Math.Min(255.0, target + 1.5);

                c.Width = target;
                c.Style.Alignment.ShrinkToFit = false;
            }
        }

        // ===== Вспомогательные парсеры/детекторы =====

        private static bool TryParseDateFlexible(string s, out DateTime parsed)
        {
            string[] fmts = {
                "yyyy-MM-dd","yyyy-MM-dd HH:mm:ss",
                "dd.MM.yyyy","dd.MM.yyyy HH:mm:ss",
                "MM/dd/yyyy","MM/dd/yyyy HH:mm:ss",
                "dd/MM/yyyy","dd/MM/yyyy HH:mm:ss",
                "yyyyMMdd","yyyyMMdd HHmmss"
            };
            if (DateTime.TryParseExact(s, fmts, null,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal, out parsed))
                return true;
            return DateTime.TryParse(s, out parsed);
        }

        private static bool LooksLikeDate(DataTable table, DataColumn col, int sample = 50)
        {
            int seen = 0, ok = 0;
            foreach (DataRow r in table.Rows)
            {
                if (seen >= sample) break;
                var v = r[col];
                if (v == DBNull.Value) continue;
                seen++;

                if (v is DateTime) { ok++; continue; }

                var s = Convert.ToString(v, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(s)) continue;

                if (TryParseDateFlexible(s, out _)) ok++;
            }
            return ok > 0;
        }

        // ===== Отладочный лист (видимый) =====

        private static void AppendDebugSheetVisible(XLWorkbook wb, string sheetName, List<ColumnPlan> plan, List<string[]> rows)
        {
            var ws = wb.Worksheets.FirstOrDefault(x => x.Name == "_ExportDebug") ?? wb.Worksheets.Add("_ExportDebug");
            ws.Visibility = XLWorksheetVisibility.VeryHidden;

            int startRow = ws.LastRowUsed()?.RowNumber() + 2 ?? 1;
            if (startRow <= 1)
            {
                ws.Cell(1, 1).Value = "Table";
                ws.Cell(1, 2).Value = "Column";
                ws.Cell(1, 3).Value = "Sys";
                ws.Cell(1, 4).Value = "DescType";
                ws.Cell(1, 5).Value = "TargetKind";
                ws.Cell(1, 6).Value = "ForcedBySTRING";
                ws.Cell(1, 7).Value = "ForcedByTableRule";
                ws.Cell(1, 8).Value = "HeuristicDate";
                ws.Row(1).Style.Font.Bold = true;
            }

            int r = startRow;
            foreach (var row in rows)
            {
                for (int c = 0; c < row.Length; c++)
                    ws.Cell(r, c + 1).Value = row[c];
                r++;
            }
        }

        // ===== Общие утилиты =====

        private static void SaveWorkbook(XLWorkbook wb, string filePath)
        {
            try
            {
                wb.SaveAs(filePath);
            }
            catch (IOException)
            {
                var dir = Path.GetDirectoryName(filePath) ?? "";
                var name = Path.GetFileNameWithoutExtension(filePath);
                var ext = Path.GetExtension(filePath);
                var alt = Path.Combine(dir, $"{name}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");
                wb.SaveAs(alt);
            }
        }

        private static string MakeSafeSheetName(string name)
        {
            var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (var ch in invalid) name = name.Replace(ch, '_');
            if (name.Length > 31) name = name.Substring(0, 31);
            return name.Trim();
        }
        private static string? TryGetDescType(DescriptorMeta? meta, string tableName, string sysOrId)
{
    if (meta == null || string.IsNullOrWhiteSpace(tableName) || string.IsNullOrWhiteSpace(sysOrId))
        return null;

    const StringComparison CI = StringComparison.OrdinalIgnoreCase;

    // meta.Fieldsets : IEnumerable
    var fsProp = meta.GetType().GetProperty("Fieldsets");
    var fieldsetsObj = fsProp?.GetValue(meta) as IEnumerable;
    if (fieldsetsObj == null) return null;

    foreach (var fs in fieldsetsObj)
    {
        if (fs == null) continue;
        var fsType = fs.GetType();

        string tbl  = fsType.GetProperty("Table")?.GetValue(fs)?.ToString() ?? "";
        string fsSn = fsType.GetProperty("SystemName")?.GetValue(fs)?.ToString() ?? "";

        if (!string.Equals(tbl,  tableName, CI) &&
            !string.Equals(fsSn, tableName, CI))
            continue;

        // fs.Fields : IEnumerable
        var fieldsObj = fsType.GetProperty("Fields")?.GetValue(fs) as IEnumerable;
        if (fieldsObj == null) continue;

        foreach (var f in fieldsObj)
        {
            if (f == null) continue;
            var ft = f.GetType();

            string sys = ft.GetProperty("SystemName")?.GetValue(f)?.ToString() ?? "";
            string id  = ft.GetProperty("Id")?.GetValue(f)?.ToString() ?? "";

            if (string.Equals(sys, sysOrId, CI) || string.Equals(id, sysOrId, CI))
            {
                return ft.GetProperty("DataType")?.GetValue(f)?.ToString();
            }
        }
    }
    return null;
}

    }
}
