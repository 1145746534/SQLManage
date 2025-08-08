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
using System.IO;
using MahApps.Metro.Controls.Dialogs;
using System.Threading;
using System.Globalization;
using Org.BouncyCastle.Asn1.X509;

namespace SQLManage.ViewModels
{
    internal class ReportManagementViewModel : BindableBase
    {
        public string Title { get; set; } = "报表管理";

        #region==============属性================


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

        private ObservableCollection<string> _items;
        /// <summary>
        /// 下拉文件框
        /// </summary>
        public ObservableCollection<string> Items
        {
            get => _items;
            set
            {
                SetProperty(ref _items, value);
            }
        }

        private string _selectedItem;
        public string SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (value == null)
                {
                    SelectPath = null;
                }
                if (result.ContainsKey(value))
                {
                    SelectPath = result[value];
                }

                SetProperty(ref _selectedItem, value);
            }
        }

        private string _selectPath;
        public string SelectPath
        {
            get => _selectPath;
            set
            {
                SetProperty(ref _selectPath, value);
            }
        }

        private Visibility _progreVisibility;
        public Visibility ProgreVisibility
        {
            get => _progreVisibility;
            set
            {
                SetProperty(ref _progreVisibility, value);
            }
        }

        private string _pickselectedItem;
        public string PickSelectedItem
        {
            get => _pickselectedItem;
            set
            {
                if (value == null)
                {
                    PickSelectPath = null;
                }
                if (result.ContainsKey(value))
                {
                    PickSelectPath = result[value];
                }

                SetProperty(ref _pickselectedItem, value);
            }
        }

        private string _pickselectPath;
        public string PickSelectPath
        {
            get => _pickselectPath;
            set
            {
                SetProperty(ref _pickselectPath, value);
            }
        }

        private Visibility _pickprogreVisibility;
        public Visibility PickProgreVisibility
        {
            get => _pickprogreVisibility;
            set
            {
                SetProperty(ref _pickprogreVisibility, value);
            }
        }


        #endregion
        #region==============命令================
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
        /// <summary>
        /// 修改数据库
        /// </summary>
        public DelegateCommand UpdataRecordCommand { get; set; }
        /// <summary>
        /// 文件挑选
        /// </summary>
        public DelegateCommand PickFileCommand { get; set; }
        #endregion

        private Dictionary<string, string> result;
        private readonly string _rootPath = @"D:\VisualDatas\HistoricalImages";

        private Timer _timer;
        //private Action<Dictionary<string, string>> _updateAction;


        private readonly IDialogCoordinator _dialogCoordinator;

        public ReportManagementViewModel(IDialogCoordinator dialogCoordinator)
        {
            this._dialogCoordinator = dialogCoordinator;

            StartDateTime = DateTime.Now.AddHours(-8);
            EndDateTime = DateTime.Now;

            DataRefreshCommand = new DelegateCommand(DataRefresh);
            DataInquireCommand = new DelegateCommand(DataInquire);
            DataStatisticsCommand = new DelegateCommand(DataStatistics);
            DataExportCommand = new DelegateCommand(DataExportAsync);
            DataExportExcelCommand = new DelegateCommand(DataExportExcel);
            UpdataRecordCommand = new DelegateCommand(UpdataRecord);
            PickFileCommand = new DelegateCommand(PickFile);

            IdentificationDatas = new ObservableCollection<Tbl_productiondatamodel>();
            StatisticsDatas = new ObservableCollection<StatisticsDataModel>();

            ProgreVisibility = Visibility.Hidden;
            PickProgreVisibility = Visibility.Hidden;
            Start();
        }

        

        /// <summary>
        /// 更新数据
        /// </summary>
        private void UpdataRecord()
        {
            if (SelectPath != null)
            {
                Console.WriteLine($"UpdataRecord");
                ProgreVisibility = Visibility.Visible;
                string[] subDirectories = Directory.GetDirectories(SelectPath);

                foreach (string subDir in subDirectories)
                {
                    string WheelPath = FileRenamer.RenameDirectory(subDir);
                    string folderName = Path.GetFileName(WheelPath); //文件夹名称                                   
                    Console.WriteLine($"\n处理子目录: {folderName}");
                    string[] strs = null;
                    if (folderName.Contains('_'))
                    {
                        strs = folderName.Split('_');
                    }
                    
                    if (strs == null || strs.Length != 2)
                    {
                        continue;
                    }
                    string model = string.Empty;
                    string style = string.Empty;
                    model = strs[0].ToUpper();
                    style = strs[1].Contains("半") ? "半成品" : "成品";

                    // 获取目录下所有文件
                    var files = Directory.GetFiles(subDir);
                    if (files.Length == 0)
                    {
                        Console.WriteLine("  -- 没有找到文件");
                        continue;
                    }
                    //  处理每个文件
                    foreach (string filePath in files)
                    {
                        // 获取文件名和扩展名
                        string fileName = Path.GetFileName(filePath);
                        bool isTrue = fileName.StartsWith(folderName); //分类是否准确
                        if (!isTrue)
                        {
                            //需要调整数据
                            UpdateModelByImageFileName(fileName, filePath, model, style);
                        }
                    }


                }
                ProgreVisibility = Visibility.Hidden;
            }

        }

        /// <summary>
        /// 挑选文件
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        private void PickFile()
        {
            if (PickSelectPath != null)
            {
                PickProgreVisibility = Visibility.Visible;

                PickProgreVisibility = Visibility.Hidden;
            }
            //DisplayText = $"文件正在挑选中......";
            //var checker = new FileNameChecker();
            //await checker.CheckAndExportFolderStructure(InputFolderPath, OutputFolderPath);
            //DisplayText = $"文件挑选完成：{InputFolderPath} -> {OutputFolderPath}";
        }

        public DateTime ConvertToDateTime(string compactTime)
        {
            if (string.IsNullOrWhiteSpace(compactTime) || compactTime.Length != 12)
            {
                throw new ArgumentException("时间格式必须为12位数字: yyMMddHHmmss");
            }

            return DateTime.ParseExact(
                compactTime,
                "yyMMddHHmmss",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None
            );
        }


        public void UpdateModelByImageFileName(string fileName, string newImagePath, string newModelName, string style)
        {
            try
            {
                //先把文件名解析
                string[] strings = fileName.Split('&');
                if (strings.Length != 2)
                {
                    return;
                }
                string compactTime = strings[1].Split('.')[0];

                DateTime result = ConvertToDateTime(compactTime);
                Console.WriteLine(result.ToString("yyyy-MM-dd HH:mm:ss"));
                DateTime startTime = result.AddMinutes(30);
                DateTime endTime = result.AddMinutes(-30);


                using (SqlSugarClient db = new SqlAccess().SystemDataAccess)
                {
                    var record = db.Queryable<Tbl_productiondatamodel>()
                                .Where(t => t.RecognitionTime > startTime && t.RecognitionTime < endTime)
                                .Where(t => t.ImagePath != null && t.ImagePath.Contains(fileName))
                                .First();
                    //// 查询包含指定文件名的记录
                    //var record = db.Queryable<Tbl_productiondatamodel>()
                    //    .Where(t => t.ImagePath.Contains(fileName))
                    //    .Where(t => t.RecognitionTime > startTime && t.RecognitionTime<endTime).Take(1).First();

                    if (record != null)
                    {
                        // 更新Model字段
                        record.Model = newModelName;
                        record.ImagePath = newImagePath;
                        record.WheelStyle = style;

                        // 更新数据库记录
                        db.Updateable(record)
                          .Where(t => t.ID == record.ID) // 确保使用主键定位记录
                          .ExecuteCommand();
                    }
                    else
                    {
                        // 可选：处理未找到记录的情况
                        //throw new Exception($"未找到包含文件名的记录: {fileName}");
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void DataRefresh()
        {

            StatisticsDataVisibility = Visibility.Collapsed;

            // 使用资源自动释放
            using (SqlSugarClient pDB = new SqlAccess().SystemDataAccess)
            {
                try
                {
                    // 直接获取最新100条记录
                    var productionList = pDB.Queryable<Tbl_productiondatamodel>()
                                          .OrderBy(x => x.ID, OrderByType.Desc)
                                          .Take(100)
                                          .ToList();

                    // 优化集合更新
                    if (IdentificationDatas == null)
                    {
                        IdentificationDatas = new ObservableCollection<Tbl_productiondatamodel>(
                            productionList.OrderByDescending(x => x.ID));
                    }
                    else
                    {
                        // 避免界面闪烁的增量更新
                        IdentificationDatas.Clear();
                        foreach (var item in productionList.OrderByDescending(x => x.ID))
                        {
                            IdentificationDatas.Add(item);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 错误处理
                    MessageBox.Show($"数据加载失败: {ex.Message}");
                }
            }

            IdentificationDataVisibility = Visibility.Visible;
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
            var productionList = pDB.Queryable<Tbl_productiondatamodel>()
                .Where(it => it.RecognitionTime > StartDateTime && it.RecognitionTime <= EndDateTime)
                .OrderBy((sc) => sc.ID, OrderByType.Desc).ToList();
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
            var productionList = pDB.Queryable<Tbl_productiondatamodel>()
                                                        .Where(it => it.RecognitionTime > StartDateTime &&it.RecognitionTime <= EndDateTime)
                                                        .ToList();
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
            var controller = await this._dialogCoordinator.ShowProgressAsync(this, "数据导出", "数据导出到本地中");
            controller.SetIndeterminate();

            //await Task.Delay(3000);       
            await Task.Run(() =>
            {

                List<Tbl_productiondatamodel> list = IdentificationDatas.ToList();

                ExportProducts("半成品", list, workShift);
                path = ExportProducts("成品", list, workShift);
            });
            await controller.CloseAsync();
            Console.WriteLine($"数据导出完成");
            //EventMessage.SystemMessageDisplay("数据导出完成", MessageType.Success);
            await Task.Delay(500);
            Process.Start("explorer.exe", path);


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

            StartDateTime = lastShiftStart;
            EndDateTime = lastShiftEnd;
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
                int setRow = 807 + appendIndex;

                int matchRow = 802;
                int macthStartCol = 1, macthEndCol = 1;

                string matchName = "班次";

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
                matchRow = 802;
                macthStartCol = macthEndCol = 3;

                matchName = "轮型";
                //setRow = 807 + appendIndex;
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
                        matchRow = 802;
                        matchName = "成品量";

                        macthStartCol = macthEndCol = 5;

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
                            matchRow = 805;
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

                    //setRow = 807 + appendIndex;  //行增加



                }

                Console.WriteLine("--------------- 统计结束 -------------------");
                appendIndex = appendIndex + 1;
            }
            ExcelHelper excelHelper = new ExcelHelper();
            string subDir = $"D:\\数据导出\\";
            string targetDir = $"{subDir}{style}\\";
            if (style == "半成品")
            {
                string path = CopyFileWithDateName("D:\\ZS\\半成品模板.xlsx", targetDir);

                excelHelper.ModifyExcelFile(exportDatas, "D:\\ZS\\半成品模板.xlsx", path);

            }
            if (style == "成品")
            {
                string path = CopyFileWithDateName("D:\\ZS\\成品模板.xlsx", targetDir);
                excelHelper.ModifyExcelFile(exportDatas, "D:\\ZS\\成品模板.xlsx", path);

            }


            exportDatas.Clear();
            exportDatas = null;
            return subDir;
        }

        public string CopyFileWithDateName(string sourceFile, string targetDir)
        {

            // 1. 校验源文件是否存在
            if (!File.Exists(sourceFile))
                throw new FileNotFoundException("源文件不存在: " + sourceFile);

            // 2. 创建目标文件夹（若不存在）
            Directory.CreateDirectory(targetDir);
            // DateTime.Now.ToString("yyyyMMdd HH-mm-ss")
            string dateStr = "";
            // 3. 生成新文件名（格式：年月日）
            if (StartDateTime != null)
            {
                dateStr = $"{StartDateTime?.ToString("yyyyMMdd HH-mm-ss")}";
            }
            if (EndDateTime != null)
            {
                dateStr = $"{dateStr}到{EndDateTime?.ToString("yyyyMMdd HH-mm-ss")}";
            }
            string style = string.Empty;
            if (sourceFile.Contains("半成品"))
            {
                style = "半成品";

            }
            else
            {
                style = "成品";
            }
            dateStr = $"{dateStr}({style}-{DateTime.Now.ToString("HH-mm-ss")})"; // 格式示例：20250618[3,7](@ref)


            string originalName = Path.GetFileNameWithoutExtension(sourceFile);
            string extension = Path.GetExtension(sourceFile);

            string newFileName = $"{dateStr}{extension}"; // 保留原名+日期后缀[3](@ref)
            string destPath = Path.Combine(targetDir, newFileName);

            // 4. 执行复制（覆盖同名文件）
            File.Copy(sourceFile, destPath, true); // true 表示覆盖[9,10](@ref)
            Console.WriteLine($"文件已复制并重命名：{destPath}");
            return destPath;


        }


        /// <summary>
        /// 分组统计（带汇总信息）
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


        public void Start()
        {
            // 立即执行一次扫描
            ScanAndUpdate();

            // 设置每天定时执行
            ResetTimer();
        }
        private void ResetTimer()
        {
            // 停止现有定时器
            _timer?.Dispose();

            // 计算到第二天凌晨的时间
            var now = DateTime.Now;
            var nextRun = now.Date.AddDays(1); // 明天凌晨
            var initialDelay = nextRun - now;

            // 创建新定时器
            _timer = new Timer(_ =>
            {
                // 在UI线程执行更新
                Application.Current.Dispatcher.Invoke(() =>
                {
                    ScanAndUpdate();
                    ResetTimer(); // 重新设置定时器
                });
            }, null, initialDelay, Timeout.InfiniteTimeSpan);
        }
        private void ScanAndUpdate()
        {
            try
            {
                result = DirectoryScanner.ScanDirectories(_rootPath);
                foreach (KeyValuePair<string, string> item in result)
                {
                    Console.WriteLine($"键：{item.Key} - 值：{item.Value}");
                }
                UpdateComboxItems(result);
            }
            catch (Exception ex)
            {
                // 处理错误，例如记录日志
                Console.WriteLine($"扫描失败: {ex.Message}");
            }
        }

        private void UpdateComboxItems(Dictionary<string, string> result)
        {
            Items?.Clear();
            Items = new ObservableCollection<string>();
            foreach (KeyValuePair<string, string> item in result)
            {
                Items.Add(item.Key);
            }

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
