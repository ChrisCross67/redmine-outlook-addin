using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Redmine.OutlookAddIn.Converters
{
    public class BooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool negate = parameter is string && (parameter as string) == "1";

            if (value is bool)
            {
                if ((bool)value == true)
                    return negate ? Visibility.Collapsed : Visibility.Visible;
            }

            return negate ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
