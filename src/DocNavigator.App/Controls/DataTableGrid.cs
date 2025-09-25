using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Data.Converters;
using Avalonia.Input;
using Avalonia.Layout;

namespace DocNavigator.App.Controls
{
    /// <summary>
    /// DataGrid для DataTable:
    /// - стабильная перестройка колонок (reset ItemsSource -> rebuild Columns -> set DefaultView);
    /// - отображение ячеек через конвертер из DataRowView/DataRow (НЕ индексеры пути);
    /// - множественный выбор строк (Ctrl/Shift);
    /// - копирование: значение / выбранные строки (TSV) / колонка;
    /// - хоткеи: Ctrl+C (значение/строки), Ctrl+Shift+C (колонка).
    /// </summary>
    public class DataTableGrid : UserControl
    {
        public static readonly StyledProperty<DataTable?> TableProperty =
            AvaloniaProperty.Register<DataTableGrid, DataTable?>(nameof(Table));

        private readonly DataGrid _grid;
        private readonly DataRowCellConverter _cellConverter = new();

        public DataTable? Table
        {
            get => GetValue(TableProperty);
            set => SetValue(TableProperty, value);
        }

        static DataTableGrid()
        {
            TableProperty.Changed.AddClassHandler<DataTableGrid>((x, _) => x.Rebuild());
        }

        public DataTableGrid()
        {
            _grid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                GridLinesVisibility = DataGridGridLinesVisibility.Horizontal,
                HeadersVisibility = DataGridHeadersVisibility.All,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                SelectionMode = DataGridSelectionMode.Extended // множественный выбор строк
            };

            // Контекст-меню
            var miCopyValue  = new MenuItem { Header = "Копировать значение\tCtrl+C" };
            miCopyValue.Click += async (_, __) => await CopyCurrentCellAsync();

            var miCopyRows   = new MenuItem { Header = "Копировать строку(и)\tCtrl+C" };
            miCopyRows.Click += async (_, __) => await CopySelectedRowsAsync(includeHeader: true);

            var miCopyColumn = new MenuItem { Header = "Копировать колонку\tCtrl+Shift+C" };
            miCopyColumn.Click += async (_, __) => await CopyCurrentColumnAsync(includeHeader: true);

            _grid.ContextMenu = new ContextMenu
            {
                ItemsSource = new MenuItem[] { miCopyValue, miCopyRows, miCopyColumn }
            };

            // Хоткеи
            _grid.KeyDown += async (_, e) =>
            {
                if (e.Key == Key.C && e.KeyModifiers == KeyModifiers.Control)
                {
                    if (_grid.SelectedItems != null && _grid.SelectedItems.Count > 0)
                        await CopySelectedRowsAsync(includeHeader: true);
                    else
                        await CopyCurrentCellAsync();
                    e.Handled = true;
                }
                else if (e.Key == Key.C && e.KeyModifiers == (KeyModifiers.Control | KeyModifiers.Shift))
                {
                    await CopyCurrentColumnAsync(includeHeader: true);
                    e.Handled = true;
                }
            };

            Content = _grid;
        }

        private void Rebuild()
        {
            // 1) Сброс источника, чтобы безопасно пересобрать схему
            _grid.ItemsSource = null;
            _grid.Columns.Clear();

            var table = Table;
            if (table == null || table.Columns.Count == 0)
                return;

            // 2) Пересоздаём колонки. Рендер через конвертер: Binding(".") + ConverterParameter = имя колонки
            foreach (DataColumn col in table.Columns)
            {
                var columnName = col.ColumnName;

                var cellTemplate = new FuncDataTemplate<object?>((_, __) =>
                {
                    var tb = new TextBlock
                    {
                        TextWrapping = Avalonia.Media.TextWrapping.NoWrap
                    };
                    tb.Bind(TextBlock.TextProperty, new Binding(".")
                    {
                        Converter          = _cellConverter,
                        ConverterParameter = columnName,
                        Mode               = BindingMode.OneWay
                    });
                    return tb;
                });

                var column = new DataGridTemplateColumn
                {
                    Header       = columnName,
                    CellTemplate = cellTemplate
                };

                _grid.Columns.Add(column);
            }

            // 3) Назначаем ItemsSource в самом конце
            _grid.ItemsSource = table.DefaultView; // элементы: DataRowView

            _grid.UpdateLayout();
            InvalidateMeasure();
            InvalidateVisual();
        }

        // ======================= Copy helpers =======================

        private async Task CopyCurrentCellAsync()
        {
            try
            {
                var rowView = _grid.SelectedItem as DataRowView
                              ?? (_grid.ItemsSource as DataView)?.Cast<DataRowView>().FirstOrDefault();
                if (rowView == null)
                {
                    await SetClipboardTextAsync(string.Empty);
                    return;
                }

                string colName =
                    (_grid.CurrentColumn?.Header?.ToString())
                    ?? _grid.Columns.FirstOrDefault()?.Header?.ToString()
                    ?? string.Empty;

                object? value = null;
                var dt = Table;
                if (dt != null && dt.Columns.Contains(colName))
                    value = rowView.Row[colName];

                await SetClipboardTextAsync(ValueToString(value));
            }
            catch
            {
                // ignore
            }
        }

        private async Task CopySelectedRowsAsync(bool includeHeader)
        {
            try
            {
                var dt = Table;
                if (dt == null) return;

                var headers = _grid.Columns.Select(c => Convert.ToString(c.Header) ?? string.Empty).ToArray();
                var sb = new StringBuilder();

                if (includeHeader)
                    sb.AppendLine(string.Join('\t', headers));

                if (_grid.SelectedItems != null && _grid.SelectedItems.Count > 0)
                {
                    foreach (var it in _grid.SelectedItems.OfType<DataRowView>())
                    {
                        var values = headers.Select(h =>
                        {
                            object? v = dt.Columns.Contains(h) ? it.Row[h] : null;
                            return ValueToString(v);
                        });
                        sb.AppendLine(string.Join('\t', values));
                    }
                }
                else
                {
                    var first = (_grid.ItemsSource as DataView)?.Cast<DataRowView>().FirstOrDefault();
                    if (first != null)
                    {
                        var values = headers.Select(h =>
                        {
                            object? v = dt.Columns.Contains(h) ? first.Row[h] : null;
                            return ValueToString(v);
                        });
                        sb.AppendLine(string.Join('\t', values));
                    }
                }

                await SetClipboardTextAsync(sb.ToString());
            }
            catch
            {
                // ignore
            }
        }

        private async Task CopyCurrentColumnAsync(bool includeHeader)
        {
            try
            {
                string colName =
                    (_grid.CurrentColumn?.Header?.ToString())
                    ?? _grid.Columns.FirstOrDefault()?.Header?.ToString()
                    ?? string.Empty;

                var dt = Table;
                if (dt == null || string.IsNullOrEmpty(colName) || !dt.Columns.Contains(colName))
                    return;

                var sb = new StringBuilder();
                if (includeHeader)
                    sb.AppendLine(colName);

                if (_grid.ItemsSource is DataView dv)
                {
                    foreach (DataRowView rowView in dv)
                    {
                        var v = rowView.Row[colName];
                        sb.AppendLine(ValueToString(v));
                    }
                }

                await SetClipboardTextAsync(sb.ToString());
            }
            catch
            {
                // ignore
            }
        }

        private static string ValueToString(object? v)
        {
            if (v == null || v == DBNull.Value) return string.Empty;
            return Convert.ToString(v, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private Task SetClipboardTextAsync(string text)
        {
            var top = TopLevel.GetTopLevel(this);
            var cb = top?.Clipboard;
            return cb != null ? cb.SetTextAsync(text) : Task.CompletedTask;
        }

        /// <summary>
        /// Конвертер, который безопасно читает значение колонки по имени из DataRowView/DataRow.
        /// </summary>
        private sealed class DataRowCellConverter : IValueConverter
        {
            public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
            {
                var col = parameter as string;
                if (string.IsNullOrEmpty(col))
                    return string.Empty;

                // DataRowView (нормальный случай в DataGrid)
                if (value is DataRowView drv)
                {
                    var t = drv.Row?.Table;
                    if (t != null && t.Columns.Contains(col))
                        return drv.Row[col];
                    return string.Empty;
                }

                // Периодически Avalonia может передать сам DataRow
                if (value is DataRow dr)
                {
                    var t = dr.Table;
                    if (t != null && t.Columns.Contains(col))
                        return dr[col];
                    return string.Empty;
                }

                return string.Empty;
            }

            public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
                throw new NotSupportedException();
        }
    }
}
