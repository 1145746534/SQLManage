using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace SQLManage.Util
{
    /// <summary>
    /// 枚举值匹配 → Visibility 转换器
    /// 当绑定值.ToString() == ConverterParameter 时返回 Visible，否则 Collapsed
    /// </summary>
    public class EnumToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return Visibility.Collapsed;
            return value.ToString() == parameter.ToString()
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
