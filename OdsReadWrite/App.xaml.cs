using System.Windows;
using System.Windows.Threading;

namespace OdsReadWrite
{
    public partial class App : Application
    {
        public App()
        {
            this.DispatcherUnhandledException += this.OnDispatcherUnhandledException;
        }

        private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString(), "Unhandled exception occured", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.None);
            e.Handled = true;
        }
    }
}
