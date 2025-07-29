using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace SQLManage.Util
{
    public class ResultConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // 将 bool 值转换为字符串
            if (value is bool boolValue)
            {
                return boolValue ? "合格" : "不合格";
            }
            if (value is string stringValue)
            {
                if (stringValue == "-1")
                {
                    return "无NG";
                }
            }
            // 如果值不是 bool 类型，返回空字符串或原始值
            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // 将字符串转换回 bool 值（通常用于双向绑定，这里可能不需要）
            if (value is string stringValue)
            {
                return stringValue == "合格";
            }
            return false;
        }
    }
}
