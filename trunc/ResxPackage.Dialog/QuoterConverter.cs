using System;
using System.Globalization;
using System.Windows.Data;

namespace ResxPackage.Dialog
{
    public class QuoterConverter: IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (double) value/4;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (double) value*4;
        }
    }
}