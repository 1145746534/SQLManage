using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using SQLManage.Models;
using SQLManage.Util;
using SqlSugar;

namespace SQLManage.ViewModels
{
    internal class ReportManagementViewModel : BindableBase
    {
        public string Title { get; set; } = "报表管理";

        #region==============日期时间相关属性================
        

        private DateTime? _startDateTime;
        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime? StartDateTime
        {
            get { return _startDateTime; }
            set
            {
                SetProperty(ref _startDateTime, value);
            }
        }


        private DateTime? _endDateTime;
        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime? EndDateTime
        {
            get { return _endDateTime; }
            set
            {
                SetProperty(ref _endDateTime, value);
            }
        }
   
        #endregion
        #region==============按钮命令相关属性================
        /// <summary>
        /// 数据刷新命令
        /// </summary>
        public DelegateCommand DataRefreshCommand { get; set; }
        /// <summary>
        /// 数据查询命令
        /// </summary>
        public DelegateCommand DataInquireCommand { get; set; }
        /// <summary>
        /// 数据统计命令
        /// </summary>
        public DelegateCommand DataStatisticsCommand { get; set; }
        /// <summary>
        /// 数据导出命令
        /// </summary>
        public DelegateCommand DataExportCommand { get; set; }

        /// <summary>
        /// 上一个班次数据导出
        /// </summary>
        public DelegateCommand DataExportExcelCommand { get; set; }
        #endregion

        private ObservableCollection<Tbl_productiondatamodel> _identificationDatas;
        /// <summary>
        /// 识别数据
        /// </summary>
        public ObservableCollection<Tbl_productiondatamodel> IdentificationDatas
        {
            get { return _identificationDatas; }
            set { SetProperty(ref _identificationDatas, value); }
        }

        private ObservableCollection<StatisticsDataModel> _statisticsDatas;
        /// <summary>
        /// 统计数据
        /// </summary>
        public ObservableCollection<StatisticsDataModel> StatisticsDatas
        {
            get { return _statisticsDatas; }
            set { SetProperty(ref _statisticsDatas, value); }
        }

        private Visibility _identificationDataVisibility;
        /// <summary>
        /// 识别数据表格显示控制
        /// </summary>
        public Visibility IdentificationDataVisibility
        {
            get { return _identificationDataVisibility; }
            set { SetProperty(ref _identificationDataVisibility, value); }
        }

        private Visibility _statisticsDataVisibility;
        /// <summary>
        /// 统计数据表格显示控制
        /// </summary>
        public Visibility StatisticsDataVisibility
        {
            get { return _statisticsDataVisibility; }
            set { SetProperty(ref _statisticsDataVisibility, value); }
        }


        public ReportManagementViewModel()
        {
            StartDateTime = DateTime.Now.AddHours(-8);
            EndDateTime = DateTime.Now;

            DataRefreshCommand = new DelegateCommand(DataRefresh);
            DataInquireCommand = new DelegateCommand(DataInquire);
            DataStatisticsCommand = new DelegateCommand(DataStatistics);
            DataExportCommand = new DelegateCommand(DataExportAsync);
            DataExportExcelCommand = new DelegateCommand(DataExportExcel);

            IdentificationDatas = new ObservableCollection<Tbl_productiondatamodel>();
            StatisticsDatas = new ObservableCollection<StatisticsDataModel>();
        }
        private void DataRefresh()
        {
            Console.WriteLine($"选择结束时间：{EndDateTime}");
            var pDB = new SqlAccess().SystemDataAccess;
            int maxId = pDB.Queryable<Tbl_productiondatamodel>().Max(x => x.ID);
            StatisticsDataVisibility = Visibility.Collapsed;
            List<Tbl_productiondatamodel> productionList = new List<Tbl_productiondatamodel>();
            if (maxId > 100)
            {
                //查询两个序号区间的数据
                productionList = pDB.Queryable<Tbl_productiondatamodel>().Where(it => SqlFunc.Between(it.ID, maxId - 100, maxId)).OrderBy((sc) => sc.ID, OrderByType.Desc).ToList();
            }
            else
            {
                productionList = pDB.Queryable<Tbl_productiondatamodel>().Where(it => SqlFunc.Between(it.ID, 0, maxId)).OrderBy((sc) => sc.ID, OrderByType.Desc).ToList();
            }
            IdentificationDatas?.Clear();
            IdentificationDatas = new ObservableCollection<Tbl_productiondatamodel>(productionList);
            IdentificationDataVisibility = Visibility.Visible;
            pDB.Close();
            pDB.Dispose();
        }
        private void DataInquire()
        {
            if (StartDateTime == null)
            {
                return;
            }
            if (EndDateTime == null)
            {
                return;
            }

            StatisticsDataVisibility = Visibility.Collapsed;
            var pDB = new SqlAccess().SystemDataAccess;
            var productionList = pDB.Queryable<Tbl_productiondatamodel>().
                Where(it => SqlFunc.Between(it.RecognitionTime, StartDateTime, EndDateTime)).
                OrderBy((sc) => sc.ID, OrderByType.Desc).ToList();
            IdentificationDatas?.Clear();
            IdentificationDatas = new ObservableCollection<Tbl_productiondatamodel>(productionList);
            IdentificationDataVisibility = Visibility.Visible;
            pDB.Close(); pDB.Dispose();

            //else 
            //    EventMessage.SystemMessageDisplay(result.Result, MessageType.Warning);
        }
        /// <summary>
        /// 数据统计
        /// </summary>
        private void DataStatistics()
        {
            if (StartDateTime == null)
            {
                return;
            }
            if (EndDateTime == null)
            {
                return;
            }
          
            
                IdentificationDataVisibility = Visibility.Hidden;
                var pDB = new SqlAccess().SystemDataAccess;
                // 从数据库读取的数据
                var productionList = pDB.Queryable<Tbl_productiondatamodel>().Where(it => SqlFunc.Between(it.RecognitionTime, StartDateTime, EndDateTime)).ToList();
                // 生成统计结果
                List<StatisticsDataModel> statistics = GenerateStatistics(productionList);

                // 输出结果
                foreach (var stat in statistics)
                {
                    Console.WriteLine($"序号: {stat.Index}, 轮型: {stat.Model}, 样式: {stat.WheelStyle}, " +
                                      $"数量: {stat.WheelCount}, 合格数: {stat.PassCount}, 主要NG: {stat.MostOfNG}");
                }
                StatisticsDatas?.Clear();
                StatisticsDatas = new ObservableCollection<StatisticsDataModel>(statistics);
                StatisticsDataVisibility = Visibility.Visible;
                //EventMessage.SystemMessageDisplay("数据统计完成", MessageType.Success);

            
            //else
            //    EventMessage.SystemMessageDisplay(result.Result, MessageType.Warning);
        }

        public List<StatisticsDataModel> GenerateStatistics(List<Tbl_productiondatamodel> productionData)
        {
            // 按Model和WheelStyle分组
            var groupedData = productionData
                .GroupBy(p => new { p.Model, p.WheelStyle })
                .Select((group, index) => new StatisticsDataModel
                {
                    Index = index + 1,
                    Model = group.Key.Model,
                    WheelStyle = group.Key.WheelStyle,
                    WheelCount = group.Count(),
                    PassCount = group.Count(p => p.Remark == "-1"),
                    MostOfNG = GetTopThreeNGs(group)
                })
                .ToList();

            return groupedData;
        }
        private string GetTopThreeNGs(IGrouping<dynamic, Tbl_productiondatamodel> group)
        {
            // 排除Remark为-1的记录（合格品）
            var ngGroups = group
                .Where(p => p.Remark != "-1")
                .GroupBy(p => p.Remark)
                .Select(g => new
                {
                    Remark = g.Key,
                    Count = g.Count()
                })
                .OrderByDescending(g => g.Count)
                .Take(3) // 取前3
                .ToList();

            if (!ngGroups.Any())
            {
                return "无NG记录";
            }

            // 拼接成字符串，格式如：A01(5),B02(3),C03(1)
            return string.Join(", ", ngGroups.Select(g => $"{g.Remark}({g.Count})"));
        }
        /// <summary>
        /// 手动导出指定日期数据
        /// </summary>
        private async void DataExportAsync()
        {
            if (IdentificationDatas.Count == 0 && StatisticsDatas.Count == 0)
            {
                //EventMessage.SystemMessageDisplay("无导出的数据，请检查！", MessageType.Warning);
                return;
            }

            //班次
            DateTime now = DateTime.Now;
            DateTime today = now.Date;
            DateTime today8 = today.AddHours(8);
            DateTime today20 = today.AddHours(20);
            string workShift = string.Empty;
            if (now >= today8 && now < today20)
            {
                // 当前是A班
                workShift = "白";
            }
            else
            {
                workShift = "晚";
            }
            string path = string.Empty;
            await Task.Run(() =>
            {

                List<Tbl_productiondatamodel> list = IdentificationDatas.ToList();

                ExportProducts("半成品", list, workShift);
                path = ExportProducts("成品", list, workShift);


                //PrintSummaryResults(summaryResults);

            });

            Console.WriteLine($"数据导出完成");
            //EventMessage.SystemMessageDisplay("数据导出完成", MessageType.Success);
            await Task.Delay(500);
            Process.Start("explorer.exe", path);

            //SaveFileDialog saveFileDialog = new SaveFileDialog
            //{
            //    Title = "请选择要导出的位置",
            //    Filter = "Excel文件(*.xls,*.xlsx)|*.xls;*.xlsx"
            //};
            //if (saveFileDialog.ShowDialog() != true) return;
            //if (saveFileDialog.FileName != "")
            //{
            //    var FileSavePath = saveFileDialog.FileName.ToString();
            //    DataTable datas = new DataTable();
            //    if (IdentificationDatas.Count > 0) datas = ExcelDataAccess.ListToDataTable(IdentificationDatas);
            //    else if (StatisticsDatas.Count > 0) datas = ExcelDataAccess.ListToDataTable(StatisticsDatas);
            //    var result = ExcelDataAccess.DataTableToExcel(datas, FileSavePath, out string exportResult);
            //    if (result) EventMessage.SystemMessageDisplay(exportResult, MessageType.Success);
            //    else EventMessage.SystemMessageDisplay(exportResult, MessageType.Error);
            //}
        }


        private async void DataExportExcel()
        {

            await AsyncExport();

            Console.WriteLine($"数据导出完成");
            //EventMessage.SystemMessageDisplay("数据导出完成", MessageType.Success);
        }

        public async Task<bool> AsyncExport()
        {
            //1.查询上一个班次的数据
            DateTime now = DateTime.Now;
            DateTime today = now.Date;
            DateTime today8 = today.AddHours(8);
            DateTime today20 = today.AddHours(20);
            string workShift = string.Empty;
            DateTime lastShiftStart, lastShiftEnd;
            if (now >= today8 && now < today20)
            {
                // 当前是A班
                lastShiftStart = today.AddDays(-1).AddHours(20); // 昨天20点
                lastShiftEnd = today8; // 今天8点
                workShift = "晚";
            }
            else
            {
                workShift = "白";
                // 当前是B班
                if (now >= today20)
                {
                    // 今天晚上20:00之后，属于今天的B班，上一个班次是今天的A班
                    lastShiftStart = today8;
                    lastShiftEnd = today20;
                }
                else
                {
                    // 当前时间小于today8，属于今天凌晨（从昨天20:00到今天8:00），上一个班次是昨天的A班
                    lastShiftStart = today.AddDays(-1).AddHours(8);
                    lastShiftEnd = today.AddDays(-1).AddHours(20);
                }
            }


            await Task.Run(() =>
            {

                var db = new SqlAccess().SystemDataAccess;
                List<Tbl_productiondatamodel> list = db.Queryable<Tbl_productiondatamodel>()
                    .Where(it => it.RecognitionTime >= lastShiftStart && it.RecognitionTime < lastShiftEnd)
                    .ToList();
                db.Close(); db.Dispose();

                // 获取所有成品
                List<Tbl_productiondatamodel> finishedProducts = list
                    .Where(p => p.WheelStyle == "成品")
                    .ToList();

                // 获取所有半成品
                List<Tbl_productiondatamodel> NotfinishedProducts = list
                    .Where(p => p.WheelStyle == "半成品")
                    .ToList();

                //ExportProducts(finishedProducts, workShift);
                ExportProducts("半成品", list, workShift);
                ExportProducts("成品", list, workShift);


                //PrintSummaryResults(summaryResults);

            });

            return true;
        }

        private string ExportProducts(string style, List<Tbl_productiondatamodel> products, string workShift)
        {

            List<Tbl_productiondatamodel> list = products
                .Where(p => p.WheelStyle == style)
                .ToList();

            List<ModelGroupSummary> summaryResults = GroupByModelThenRemarkWithSummary(list);
            //每一个单元格都需要往队列里面添数据
            Queue<ExportDataModel> exportDatas = new Queue<ExportDataModel>();
            int appendIndex = 0;
            foreach (var modelSummary in summaryResults)
            {
                // 每次循环写入一行数据

                int matchRow = 760;
                int macthStartCol = 1;
                int macthEndCol = 1;
                string matchName = "班次";
                int setRow = 765 + appendIndex;
                object setValue = workShift;
                //班次
                exportDatas.Enqueue(new ExportDataModel()
                {
                    MatchRow = matchRow,
                    MatchName = matchName,
                    SettingRow = setRow,
                    SettingValue = setValue,
                    MatchStartCol = macthStartCol,
                    MatchEndCol = macthEndCol,
                });
                //单元


                //轮形
                matchRow = 760;
                macthStartCol = 3;
                macthEndCol = 3;
                matchName = "轮型";
                setRow = 765 + appendIndex;
                setValue = modelSummary.Model;

                exportDatas.Enqueue(new ExportDataModel()
                {
                    MatchRow = matchRow,
                    MatchName = matchName,
                    SettingRow = setRow,
                    SettingValue = setValue,
                    MatchStartCol = macthStartCol,
                    MatchEndCol = macthEndCol
                });



                Console.WriteLine($"型号: {modelSummary.Model} - {style}");
                Console.WriteLine($"  总记录数: {modelSummary.TotalCount}");


                foreach (var remarkGroup in modelSummary.RemarkGroups)
                {
                    double percentage = (double)remarkGroup.Count / modelSummary.TotalCount * 100;
                    Console.WriteLine($"    - 备注: {remarkGroup.Remark}, 数量: {remarkGroup.Count} ({percentage:F1}%)");



                    if (string.IsNullOrEmpty(remarkGroup.Remark) || remarkGroup.Remark == "-1") //为空或者-1就是OK的产品
                    {

                        //OK量
                        matchRow = 760;
                        matchName = "成品量";

                        macthStartCol = 5;
                        macthEndCol = 5;
                        setValue = remarkGroup.Count;

                        exportDatas.Enqueue(new ExportDataModel()
                        {
                            MatchRow = matchRow,
                            MatchStartCol = macthStartCol,
                            MatchEndCol = macthEndCol,
                            MatchName = matchName,
                            SettingRow = setRow,
                            SettingValue = setValue
                        });
                    }
                    else
                    {
                        //NG的 线上提报  or 平板提报
                        foreach (var reportWayGroup in remarkGroup.ReportWayGroups)
                        {
                            Console.WriteLine($"  │   ├─ [报告方式] {reportWayGroup.ReportWay} (数量: {reportWayGroup.Count})");

                            string result = remarkGroup.Remark.PadLeft(2, '0');
                            matchRow = 763;
                            matchName = $"5{result}";
                            if (reportWayGroup.ReportWay == "线上")
                            {
                                macthStartCol = 24; //X
                                macthEndCol = 128; //DX
                            }
                            if (reportWayGroup.ReportWay == "平板")
                            {

                                macthStartCol = 129; //DY 
                                macthEndCol = 242; //IH

                            }
                            setValue = reportWayGroup.Count;

                            exportDatas.Enqueue(new ExportDataModel()
                            {
                                MatchRow = matchRow,
                                MatchName = matchName,
                                MatchStartCol = macthStartCol,
                                MatchEndCol = macthEndCol,
                                SettingRow = setRow,
                                SettingValue = setValue
                            });
                        }


                    }

                    setRow = 765 + appendIndex;  //行增加



                }

                Console.WriteLine("--------------- 统计结束 -------------------");
                appendIndex = appendIndex + 1;
            }
            ExcelHelper excelHelper = new ExcelHelper();
            string subDir = $"D:\\数据导出\\";
            string targetDir = $"{subDir}{style}\\";
            excelHelper.ModifyExcelFile(exportDatas, "D:\\ZS\\数据导出.xlsx", targetDir);
            exportDatas.Clear();
            exportDatas = null;
            return subDir;
        }



        /// <summary>
        /// 双层分组统计（带汇总信息）
        /// </summary>
        public List<ModelGroupSummary> GroupByModelThenRemarkWithSummary(
            List<Tbl_productiondatamodel> dataList)
        {
            // 先按Model分组
            return dataList
                .GroupBy(item => item.Model ?? "无型号")
                .Select(modelGroup => new ModelGroupSummary
                {
                    Model = modelGroup.Key,
                    TotalCount = modelGroup.Count(),
                    RemarkGroups = modelGroup
                        .GroupBy(item => item.Remark ?? "无备注")
                        .Select(remarkGroup => new RemarkGroup
                        {
                            Remark = remarkGroup.Key,
                            Count = remarkGroup.Count(),
                            // 新增：在 RemarkGroup 下再按 ReportWay 分组
                            ReportWayGroups = remarkGroup
                                .GroupBy(item => item.ReportWay ?? "无报告方式")
                                .Select(reportWayGroup => new ReportWayGroup
                                {
                                    ReportWay = reportWayGroup.Key,
                                    Count = reportWayGroup.Count()
                                })
                                .OrderBy(rwg => rwg.ReportWay)
                                .ToList()
                        })
                        .OrderBy(rg => rg.Remark)
                        .ToList()
                })
                .OrderBy(summary => summary.Model)
                .ToList();
        }

        private static void PrintSummaryResults(List<ModelGroupSummary> summaries)
        {
            //每一个单元格都需要往队列里面添数据
            Queue<ExportDataModel> exportDatas = new Queue<ExportDataModel>();
            foreach (var modelSummary in summaries)
            {
                int matchRow = 760;
                string matchName = "班次";
                int setRow = 765;
                string setValue = "";
                Console.WriteLine($"型号: {modelSummary.Model}");
                Console.WriteLine($"  总记录数: {modelSummary.TotalCount}");


                foreach (var remarkGroup in modelSummary.RemarkGroups)
                {
                    double percentage = (double)remarkGroup.Count / modelSummary.TotalCount * 100;
                    Console.WriteLine($"    - 备注: {remarkGroup.Remark}, 数量: {remarkGroup.Count} ({percentage:F1}%)");

                }

                Console.WriteLine("----------------------------------");
            }
        }

        /// <summary>
        /// 检查输入的日期时间
        /// </summary>
        /// <param name="startDate">起始日期</param>
        /// <param name="startHour">起始小时</param>
        /// <param name="startMinute">起始分钟</param>
        /// <param name="endDate">结束日期</param>
        /// <param name="endHour">结束小时</param>
        /// <param name="endMinute">结束分钟</param>
        /// <param name="result">检查结果</param>
        /// <returns>True为无错误</returns>
        private bool JudgmentInputsDateTime(DateTime startDate, string startHour, string startMinute, DateTime endDate, string endHour, string endMinute, out GetDateTimeModel result)
        {
            GetDateTimeModel dateTime = new GetDateTimeModel();
            if (!int.TryParse(startHour, out int sh) || sh < 0 || sh > 24)
            {
                dateTime.Result = "起始小时输入错误，请检查！";
                result = dateTime;
                return false;
            }
            if (!int.TryParse(startMinute, out int sm) || sm < 0 || sm > 59)
            {
                dateTime.Result = "起始分钟输入错误，请检查！";
                result = dateTime;
                return false;
            }
            if (!int.TryParse(endHour, out int eh) || eh < 0 || eh > 24)
            {
                dateTime.Result = "结束小时输入错误，请检查！";
                result = dateTime;
                return false;
            }
            if (!int.TryParse(endMinute, out int em) || em < 0 || em > 59)
            {
                dateTime.Result = "结束分钟输入错误，请检查！";
                result = dateTime;
                return false;
            }
            GenDateTime(startDate, startHour, startMinute, out DateTime startDateTime);
            GenDateTime(endDate, endHour, endMinute, out DateTime endDateTime);
            if (startDateTime >= endDateTime)
            {
                dateTime.Result = "起始日期大于或等于结束日期，请检查！";
                result = dateTime;
                return false;
            }
            dateTime.StartDateTime = startDateTime;
            dateTime.EndDateTime = endDateTime;
            dateTime.Result = "OK";
            result = dateTime;
            return true;
        }

        /// <summary>
        /// 生成日期时间
        /// </summary>
        /// <param name="date">日期</param>
        /// <param name="hour">小时</param>
        /// <param name="minute">分钟</param>
        /// <param name="dateTime">生成的日期时间</param>
        /// <returns>True为生成成功</returns>
        private bool GenDateTime(DateTime date, string hour, string minute, out DateTime dateTime)
        {
            string strDate = date.ToString().Replace("0:00:00", "");
            string strDateTime = strDate + hour + ":" + minute + ":00";
            return DateTime.TryParse(strDateTime, out dateTime);
        }
    }
    public class ModelGroupSummary
    {
        public string Model { get; set; }
        public int TotalCount { get; set; }        // 该型号总条数
        public List<RemarkGroup> RemarkGroups { get; set; } // 该型号下的备注分组
    }

    public class RemarkGroup
    {
        /// <summary>
        /// NG编号
        /// </summary>
        public string Remark { get; set; }


        public int Count { get; set; }

        public List<ReportWayGroup> ReportWayGroups { get; set; }

    }

    public class ReportWayGroup
    {
        public string ReportWay { get; set; }
        public int Count { get; set; }
    }
}
