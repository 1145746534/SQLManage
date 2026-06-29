using System.Windows;
using System.Windows.Controls;
using SQLManage.ViewModels;

namespace SQLManage.Util
{
    /// <summary>
    /// 根据 DataViewType 切换不同的 DataGrid 模板
    /// </summary>
    public class DataViewTemplateSelector : DataTemplateSelector
    {
        public DataTemplate IdentificationTemplate { get; set; }
        public DataTemplate StatisticsTemplate { get; set; }
        public DataTemplate PaintingTotalTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (item is ReportManagementViewModel.DataViewType viewType)
            {
                switch (viewType)
                {
                    case ReportManagementViewModel.DataViewType.Identification:
                        return IdentificationTemplate;
                    case ReportManagementViewModel.DataViewType.Statistics:
                        return StatisticsTemplate;
                    case ReportManagementViewModel.DataViewType.PaintingTotal:
                        return PaintingTotalTemplate;
                }
            }
            return null;
        }
    }
}
