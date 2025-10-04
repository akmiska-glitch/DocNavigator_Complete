using System;
using System.Collections.Generic;
using System.Data;
using DocNavigator.App.Models;

namespace DocNavigator.App.Services.Metadata
{
    public static class FieldLocalizer
    {
        /// <summary>
        /// Применяет русские подписи к колонкам согласно DescriptorMeta.ColumnCaptionsById.
        /// Сопоставление по field/@id без учета регистра.
        /// </summary>
        public static void ApplyDisplayNames(DataTable table, DescriptorMeta? meta)
        {
            if (table == null || table.Columns.Count == 0 || meta == null)
                return;

            var src = meta.ColumnCaptionsById;
            if (src == null || src.Count == 0)
                return;

            // Создаем регистронезависимую карту
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in src)
            {
                if (!string.IsNullOrWhiteSpace(kv.Key) && !string.IsNullOrWhiteSpace(kv.Value))
                    map[kv.Key] = kv.Value;
            }
            if (map.Count == 0) return;

            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (DataColumn col in table.Columns)
            {
                var key = col.ColumnName;
                if (map.TryGetValue(key, out var ru) && !string.IsNullOrWhiteSpace(ru))
                {
                    var unique = MakeUnique(ru, used);
                    col.ColumnName = unique;
                }
                else
                {
                    used.Add(col.ColumnName);
                }
            }

        }

        private static string MakeUnique(string baseName, HashSet<string> used)
        {
            if (!used.Contains(baseName))
            {
                used.Add(baseName);
                return baseName;
            }
            int i = 2;
            while (true)
            {
                var cand = $"{baseName} ({i})";
                if (!used.Contains(cand))
                {
                    used.Add(cand);
                    return cand;
                }
                i++;
            }
        }
        
    }
    
}