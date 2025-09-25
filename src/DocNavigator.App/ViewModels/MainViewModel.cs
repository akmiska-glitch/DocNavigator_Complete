using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using DocNavigator.App.Models;
using DocNavigator.App.Services.Data;
using DocNavigator.App.Services.Export;
using DocNavigator.App.Services.Metadata;
using DocNavigator.App.Services.Profiles;

namespace DocNavigator.App.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));

        private readonly ProfileService _profileService;
        private readonly DescParser _descParser;
        private readonly IRemoteDescProvider _remoteDesc;

        public ObservableCollection<DbProfile> Profiles { get; } = new();
        private DbProfile? _selectedProfile;
        public DbProfile? SelectedProfile
        {
            get => _selectedProfile;
            set { _selectedProfile = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanQuery)); }
        }

        public ObservableCollection<TableNodeViewModel> Tables { get; } = new();
        private TableNodeViewModel? _selectedTable;
        public TableNodeViewModel? SelectedTable
        {
            get => _selectedTable;
            set
            {
                _selectedTable = value; OnPropertyChanged();
                _ = OpenSelectedAsync();
            }
        }

        public ObservableCollection<TableTabViewModel> OpenTabs { get; } = new();
        private TableTabViewModel? _activeTab;
        public TableTabViewModel? ActiveTab
        {
            get => _activeTab;
            set { _activeTab = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasActiveTab)); }
        }
        public bool HasActiveTab => ActiveTab != null;
        public bool HasAnyTabs => OpenTabs.Count > 0;

        public bool CanExportAll => Tables.Any(t => t.HasRows);

        private bool _useRussianCaptions = false;
        public bool UseRussianCaptions
        {
            get => _useRussianCaptions;
            set
            {
                if (_useRussianCaptions == value) return;
                _useRussianCaptions = value; OnPropertyChanged();
                RebuildAllOpenTabs();
            }
        }

        private bool _hideEmptyColumns = false;
        public bool HideEmptyColumns
        {
            get => _hideEmptyColumns;
            set
            {
                if (_hideEmptyColumns == value) return;
                _hideEmptyColumns = value; OnPropertyChanged();
                RebuildAllOpenTabs();
            }
        }

        private string _guidInput = string.Empty;
        public string GuidInput
        {
            get => _guidInput;
            set { _guidInput = value; OnPropertyChanged(); OnPropertyChanged(nameof(CanQuery)); }
        }

        public bool CanQuery => SelectedProfile != null && !string.IsNullOrWhiteSpace(GuidInput);

        public ICommand QueryByGuidCommand { get; }
        public ICommand CloseTabCommand { get; }
        public ICommand ExportActiveTabToExcelCommand { get; }
        public ICommand ExportAllTabsToExcelCommand { get; }
        public ICommand ExportActiveTabToCsvRawCommand { get; }

        private IDocumentRepository? _lastRepo;
        private DocRow? _lastDoc;
        private DescriptorMeta? _currentDescMeta;

        private readonly Dictionary<string, string> _tableCaptions = new(StringComparer.OrdinalIgnoreCase);

        public MainViewModel(ProfileService profileService, DescParser descParser, IRemoteDescProvider remoteDesc)
        {
            _profileService = profileService;
            _descParser = descParser;
            _remoteDesc = remoteDesc;

            foreach (var p in _profileService.LoadAll())
                Profiles.Add(p);
            SelectedProfile = Profiles.FirstOrDefault();

            OpenTabs.CollectionChanged += (_, __) => OnPropertyChanged(nameof(HasAnyTabs));
            Tables.CollectionChanged += (_, __) => OnPropertyChanged(nameof(CanExportAll));

            QueryByGuidCommand = new RelayCommand(async _ => await QueryByGuidAsync());
            CloseTabCommand = new RelayCommand(tab =>
            {
                if (tab is TableTabViewModel t)
                    OpenTabs.Remove(t);
            });

            ExportActiveTabToExcelCommand = new RelayCommand(async _ => await ExportActiveToExcelAsync());
            ExportAllTabsToExcelCommand   = new RelayCommand(async _ => await ExportAllToExcelAsync());
            ExportActiveTabToCsvRawCommand= new RelayCommand(async _ => await ExportActiveToCsvRawAsync());
        }

        private async Task QueryByGuidAsync()
        {
            if (!CanQuery) return;
            if (SelectedProfile == null) return;
            if (!Guid.TryParse(GuidInput, out var g)) return;

            OpenTabs.Clear();
            ActiveTab = null;
            Tables.Clear();
            _tableCaptions.Clear();
            _currentDescMeta = null;

            IDocumentRepository repo = SelectedProfile.Kind.ToLowerInvariant() switch
            {
                "postgres" => new PostgresDocumentRepository(SelectedProfile),
                "oracle"   => new OracleDocumentRepository(SelectedProfile),
                _ => throw new NotSupportedException($"Unknown DB kind: {SelectedProfile.Kind}")
            };

            var doc = await repo.FindDocByGuidAsync(g, CancellationToken.None);
            if (doc == null) return;

            _lastRepo = repo;
            _lastDoc  = doc;

            var dtype = await repo.FindDocTypeAsync(doc.DocTypeId, CancellationToken.None);

            var candidates = new List<string> { "doc", "routecontext" };

            if (dtype != null)
            {
                try
                {
                    var pair = await repo.FindDoctypeAndServiceAsync(doc.DocTypeId, CancellationToken.None);
                    if (pair != null)
                    {
                        var xml = await _remoteDesc.GetByCodesAsync(SelectedProfile!, pair.Value.ServiceCode, pair.Value.DoctypeCode, CancellationToken.None);
                        _currentDescMeta = _descParser.ParseFromText(xml);
                    }
                }
                catch {}

                if (_currentDescMeta == null)
                    _currentDescMeta = _descParser.Load(dtype.SystemName + ".desc");

                if (_currentDescMeta != null)
                {
                    if (!string.IsNullOrWhiteSpace(_currentDescMeta.ContentTable))
                        candidates.Add(_currentDescMeta.ContentTable);

                    foreach (var t in _currentDescMeta.FieldsetTables)
                        if (!string.IsNullOrWhiteSpace(t))
                            candidates.Add(t);

                    foreach (var kv in _currentDescMeta.TableCaptions)
                        _tableCaptions[kv.Key] = kv.Value;
                }
            }

            candidates = candidates
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var temp = new List<TableNodeViewModel>();
            var fsSet = new HashSet<string>(
                _currentDescMeta?.FieldsetTables ?? Enumerable.Empty<string>()
                , StringComparer.OrdinalIgnoreCase);

            foreach (var tName in candidates)
            {
                var hasRows = await repo.HasRowsAsync(tName, doc.DocId, CancellationToken.None);
                var title = MakeTitleForNav(tName);
                temp.Add(new TableNodeViewModel(
                    name: tName,
                    title: title,
                    hasRows: hasRows,
                    openAction: async () => await OpenTableAsync(tName)
                ));
            }

            var nonFs = temp.Where(x => !fsSet.Contains(x.Name)).ToList();
            var fs    = temp.Where(x =>  fsSet.Contains(x.Name))
                            .OrderByDescending(x => x.HasRows)
                            .ThenBy(x => x.Title, StringComparer.OrdinalIgnoreCase)
                            .ToList();

            Tables.Clear();
            foreach (var x in nonFs) Tables.Add(x);
            foreach (var x in fs)    Tables.Add(x);
        }

        private async Task OpenSelectedAsync()
        {
            var item = SelectedTable;
            if (item == null) return;
            await OpenTableAsync(item.Name);
        }

        private async Task OpenTableAsync(string tableName)
        {
            if (_lastRepo == null || _lastDoc == null) return;

            var existing = OpenTabs.FirstOrDefault(t => string.Equals(t.TabKey, tableName, StringComparison.OrdinalIgnoreCase));
            if (existing != null) { ActiveTab = existing; return; }

            if (!await _lastRepo.HasRowsAsync(tableName, _lastDoc.DocId, CancellationToken.None))
                return;

            var src = await _lastRepo.ReadTableAsync(tableName, _lastDoc.DocId, CancellationToken.None);
            src.TableName = tableName;

            var header = MakeTitleForNav(tableName);
            var tab = new TableTabViewModel(tableName, header, src);

            var view = BuildView(src);
            if (UseRussianCaptions)
                FieldLocalizer.ApplyDisplayNames(view, _currentDescMeta);

            tab.SetTable(view);
            OpenTabs.Add(tab);
            ActiveTab = tab;
        }

        private DataTable BuildView(DataTable src)
        {
            var copy = src.Copy();
            copy.TableName = src.TableName;

            if (HideEmptyColumns)
                copy = FilterEmptyColumns(copy);

            return copy;
        }

        private void RebuildAllOpenTabs()
        {
            foreach (var tab in OpenTabs)
            {
                var rebuilt = BuildView(tab.SourceTable);
                if (UseRussianCaptions)
                    FieldLocalizer.ApplyDisplayNames(rebuilt, _currentDescMeta);

                tab.SetTable(rebuilt);
            }
            OnPropertyChanged(nameof(OpenTabs));
            if (ActiveTab != null) OnPropertyChanged(nameof(ActiveTab));
        }

        private string MakeTitleForNav(string systemName)
        {
            var hasRu = _tableCaptions.TryGetValue(systemName, out var ru) && !string.IsNullOrWhiteSpace(ru);
            var baseTitle = hasRu ? $"{systemName} ({ru})" : systemName;
            return baseTitle;
        }

        private async Task ExportActiveToExcelAsync()
        {
            if (!HasActiveTab) return;

            var sfd = new SaveFileDialog
            {
                Title = "Сохранить вкладку в Excel",
                InitialFileName = $"{SafeFileName(ActiveTab!.Header)}.xlsx",
                Filters = new() { new FileDialogFilter { Name = "Excel", Extensions = { "xlsx" } } }
            };

            var owner = (Application.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;
            var path = await sfd.ShowAsync(owner);
            if (string.IsNullOrWhiteSpace(path)) return;

            try
            {
                ExcelExporter.SaveSingleTable(ActiveTab!.Table, ActiveTab!.Header, path);
            }
            catch (IOException ioex)
            {
                await ShowErrorAsync(ioex.Message);
            }
            catch (Exception ex)
            {
                await ShowErrorAsync("Ошибка при сохранении Excel: " + ex.Message);
            }
        }

        private async Task ExportAllToExcelAsync()
        {
            if (_lastRepo == null || _lastDoc == null || Tables.Count == 0)
                return;

            var items = new List<(string Sheet, DataTable Table)>();
            foreach (var node in Tables)
            {
                if (!node.HasRows) continue;

                var dt = await _lastRepo.ReadTableAsync(node.Name, _lastDoc.DocId, CancellationToken.None);
                if (dt.Rows.Count == 0) continue;

                var sheet = node.Title;
                items.Add((sheet, dt));
            }

            if (items.Count == 0) return;

            var sfd = new SaveFileDialog
            {
                Title = "Сохранить все таблицы (XLSX)",
                InitialFileName = $"{SafeFileName(GuidInput)}.xlsx",
                Filters = new() { new FileDialogFilter { Name = "Excel", Extensions = { "xlsx" } } }
            };

            var owner = (Application.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;
            var path = await sfd.ShowAsync(owner);
            if (string.IsNullOrWhiteSpace(path)) return;

            try
            {
                ExcelExporter.SaveMultipleTables(items.ToArray(), path);
            }
            catch (IOException ioex)
            {
                await ShowErrorAsync(ioex.Message);
            }
            catch (Exception ex)
            {
                await ShowErrorAsync("Ошибка при сохранении Excel: " + ex.Message);
            }
        }

        private async Task ExportActiveToCsvRawAsync()
        {
            if (!HasActiveTab) return;

            if (_lastRepo == null || _lastDoc == null)
            {
                var ok = await AskUserAsync("Нет контекста последнего запроса к БД.\nЭкспортировать текущую вкладку в CSV как есть?");
                if (!ok) return;

                var sfdFallback = new SaveFileDialog
                {
                    Title = "Сохранить CSV (текущая вкладка)",
                    InitialFileName = $"{SafeFileName(ActiveTab!.Header)}.csv",
                    Filters = new() { new FileDialogFilter { Name = "CSV", Extensions = { "csv" } } }
                };
                var ownerFallback = (Application.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;
                var pathFallback = await sfdFallback.ShowAsync(ownerFallback);
                if (string.IsNullOrWhiteSpace(pathFallback)) return;

                try
                {
                    CsvExporter.Save(ActiveTab!.Table, pathFallback, ',');
                }
                catch (Exception ex)
                {
                    await ShowErrorAsync("Ошибка при сохранении CSV: " + ex.Message);
                }
                return;
            }

            var raw = await _lastRepo.ReadTableAsync(ActiveTab!.TabKey, _lastDoc.DocId, CancellationToken.None);

            var sfd = new SaveFileDialog
            {
                Title = "Сохранить CSV (сырые данные)",
                InitialFileName = $"{SafeFileName(ActiveTab!.Header)}.csv",
                Filters = new() { new FileDialogFilter { Name = "CSV", Extensions = { "csv" } } }
            };

            var owner = (Application.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;
            var path = await sfd.ShowAsync(owner);
            if (string.IsNullOrWhiteSpace(path)) return;

            try
            {
                CsvExporter.Save(raw, path, ',');
            }
            catch (Exception ex)
            {
                await ShowErrorAsync("Ошибка при сохранении CSV: " + ex.Message);
            }
        }

        private static DataTable FilterEmptyColumns(DataTable source)
        {
            bool IsEmptyColumn(DataColumn col)
            {
                foreach (DataRow row in source.Rows)
                {
                    var v = row[col];
                    if (v != null && v != DBNull.Value && !string.IsNullOrWhiteSpace(Convert.ToString(v)))
                        return false;
                }
                return true;
            }

            var cols = source.Columns.Cast<DataColumn>()
                .Where(c => !IsEmptyColumn(c))
                .ToList();

            if (cols.Count == 0)
                return source;

            var clone = new DataTable(source.TableName);
            foreach (var c in cols)
                clone.Columns.Add(c.ColumnName, c.DataType);

            foreach (DataRow r in source.Rows)
            {
                var nr = clone.NewRow();
                foreach (var c in cols)
                    nr[c.ColumnName] = r[c] ?? DBNull.Value;
                clone.Rows.Add(nr);
            }
            return clone;
        }

        private static string SafeFileName(string s)
        {
            var invalid = System.IO.Path.GetInvalidFileNameChars();
            foreach (var ch in invalid) s = s.Replace(ch, '_');
            return s;
        }

        private async Task<bool> AskUserAsync(string message)
        {
            var owner = (Application.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;
            if (owner == null) return false;

            var dlg = new Window
            {
                Title = "Подтверждение",
                Width = 460,
                Height = 200,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Content = new StackPanel
                {
                    Margin = new Thickness(16),
                    Children =
                    {
                        new TextBlock { Text = message, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                        new StackPanel
                        {
                            Orientation = Avalonia.Layout.Orientation.Horizontal,
                            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
                            Margin = new Thickness(0,12,0,0),
                            Children =
                            {
                                new Button { Content = "Отмена", IsCancel = true, Margin = new Thickness(6,0) },
                                new Button { Content = "Ок", IsDefault = true, Margin = new Thickness(6,0) }
                            }
                        }
                    }
                }
            };

            bool result = false;
            ((dlg.Content as StackPanel)!.Children[1] as StackPanel)!.Children[0]
                .AddHandler(Button.ClickEvent, (_, __) => { dlg.Close(); });
            ((dlg.Content as StackPanel)!.Children[1] as StackPanel)!.Children[1]
                .AddHandler(Button.ClickEvent, (_, __) => { result = true; dlg.Close(); });

            await dlg.ShowDialog(owner);
            return result;
        }

        private async Task ShowErrorAsync(string message)
        {
            var owner = (Application.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;
            if (owner == null) return;

            var dlg = new Window
            {
                Title = "Ошибка",
                Width = 520,
                Height = 220,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Content = new StackPanel
                {
                    Margin = new Thickness(16),
                    Children =
                    {
                        new TextBlock { Text = message, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                        new StackPanel
                        {
                            Orientation = Avalonia.Layout.Orientation.Horizontal,
                            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
                            Margin = new Thickness(0,12,0,0),
                            Children = { new Button { Content = "Ок", IsDefault = true, IsCancel = true, Margin = new Thickness(6,0) } }
                        }
                    }
                }
            };

            ((dlg.Content as StackPanel)!.Children[1] as StackPanel)!.Children[0]
                .AddHandler(Button.ClickEvent, (_, __) => { dlg.Close(); });

            await dlg.ShowDialog(owner);
        }
    }
}