using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using DocumentDownloader.Models;
using DocumentDownloader.ViewModels;

namespace DocumentDownloader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel vm && sender is PasswordBox passwordBox)
            {
                vm.Password = passwordBox.Password;
            }
        }
    }

    /// <summary>
    /// Converts LogLevel enum to brush color for UI display
    /// </summary>
    public class LogLevelToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is LogLevel level)
            {
                return level switch
                {
                    LogLevel.Success => new SolidColorBrush(Color.FromRgb(16, 124, 16)),  // Green
                    LogLevel.Warning => new SolidColorBrush(Color.FromRgb(255, 140, 0)),  // Orange
                    LogLevel.Error => new SolidColorBrush(Color.FromRgb(216, 59, 1)),     // Red
                    _ => new SolidColorBrush(Color.FromRgb(0, 120, 212))                   // Blue (Info)
                };
            }
            return new SolidColorBrush(Color.FromRgb(0, 120, 212));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Converts boolean to Visibility
    /// </summary>
    public class BooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}