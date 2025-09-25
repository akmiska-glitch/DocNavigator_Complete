using System.ComponentModel;
using System.Data;
using System.Runtime.CompilerServices;

namespace DocNavigator.App.ViewModels
{
    public class TableTabViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));

        /// <summary>Ключ вкладки — системное имя таблицы. Используется для недопущения дублей.</summary>
        public string TabKey { get; }

        private string _header;
        public string Header
        {
            get => _header;
            set { _header = value; OnPropertyChanged(); }
        }

        /// <summary>Оригинальная (сырая) таблица, как вернулась из БД.</summary>
        public DataTable SourceTable { get; }

        /// <summary>Текущая отображаемая таблица (с учётом флагов и локализаций).</summary>
        private DataTable _table;
        public DataTable Table
        {
            get => _table;
            private set { _table = value; OnPropertyChanged(); }
        }

        public TableTabViewModel(string tabKey, string header, DataTable source)
        {
            TabKey = tabKey;
            _header = header;
            SourceTable = source;
            _table = source;
        }

        public void SetTable(DataTable dt) => Table = dt;
    }
}