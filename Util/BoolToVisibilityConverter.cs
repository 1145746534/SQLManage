using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows;

namespace SQLManage.Util
{
    /// <summary>
    /// Bool 值转换为 Visibility，支持 ConverterParameter="Inverted" 反转
    /// </summary>
    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool isInverted = parameter is string str && str.Equals("Inverted", StringComparison.OrdinalIgnoreCase);

            if (value is bool boolVal)
            {
                boolVal = isInverted ? !boolVal : boolVal;
                return boolVal ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool isInverted = parameter is string str && str.Equals("Inverted", StringComparison.OrdinalIgnoreCase);

            if (value is Visibility vis)
            {
                bool result = vis == Visibility.Visible;
                return isInverted ? !result : result;
            }
            return false;
        }
    }

}
