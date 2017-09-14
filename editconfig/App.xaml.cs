using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace editconfig
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                MessageBox.Show(e.ExceptionObject.ToString().TrimStart(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            };
        }
    }
}
