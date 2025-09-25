using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using DocNavigator.App.Views;
using DocNavigator.App.ViewModels;
using DocNavigator.App.Services.Profiles;
using DocNavigator.App.Services.Metadata;
using System;

namespace DocNavigator.App
{
    public partial class App : Application
    {
        public override void Initialize()
        {
            AvaloniaXamlLoader.Load(this);
        }

        public override void OnFrameworkInitializationCompleted()
        {
            var profileService = new ProfileService("Config/profiles");
            var descParser = new DescParser("Metadata/descriptors");
            var remoteDesc = new RemoteDescProvider(timeout: TimeSpan.FromSeconds(8)); // кэш на сессию

            if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                var vm = new MainViewModel(profileService, descParser, remoteDesc);
                desktop.MainWindow = new MainWindow { DataContext = vm };
            }
            base.OnFrameworkInitializationCompleted();
        }
    }
}
