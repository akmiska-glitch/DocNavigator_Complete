using System.Data;

namespace DocNavigator.App.Controls
{
    /// <summary>
    /// Обёртка над DataRow с индексатором по имени колонки.
    /// Позволяет удобно биндиться в Avalonia: Binding="[{ColumnName}]".
    /// </summary>
    public sealed class DynamicRow
    {
        private readonly DataRow _row;
        public DynamicRow(DataRow row) => _row = row;

        public object? this[string columnName] => _row[columnName] is DBNull ? null : _row[columnName];
    }
}
