using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DocNavigator.App.ViewModels
{
    public class TableNodeViewModel
    {
        public string Name { get; }
        public string Title { get; }
        public bool HasRows { get; }
        public ICommand OpenCommand { get; }
        public TableNodeViewModel(string name, string title, bool hasRows, Func<Task> openAction)
        {
            Name = name; Title = title; HasRows = hasRows;
            OpenCommand = new RelayCommand(async _ => await openAction());
        }
    }
}
