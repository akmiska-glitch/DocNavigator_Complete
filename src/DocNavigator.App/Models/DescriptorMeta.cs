using System.Collections.Generic;

namespace DocNavigator.App.Models
{
    /// <summary>
    /// Метаданные из .desc файла.
    /// </summary>
    public class DescriptorMeta
    {
        /// <summary>Основная таблица (<content table="..."/>)</summary>
        public string? ContentTable { get; set; }

        /// <summary>Таблицы из fieldset-def/nested-fieldset.</summary>
        public List<string> FieldsetTables { get; set; } = new();

        /// <summary>Подписи таблиц (tableName → caption).</summary>
        public Dictionary<string, string> TableCaptions { get; set; } = new();

        /// <summary>Поля по system name (если нужно где-то ещё).</summary>
        public Dictionary<string, FieldMeta> FieldsBySystemName { get; set; } = new();

        /// <summary>
        /// Русские подписи по ИДЕНТИФИКАТОРУ КОЛОНКИ (field @id → RU).
        /// Именно этот словарь используется для переименования колонок грида.
        /// </summary>
        public Dictionary<string, string> ColumnCaptionsById { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public object Fieldsets { get; internal set; }
    }

    
}
