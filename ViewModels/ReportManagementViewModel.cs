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
        /// <summary>
        /// 数据视图类型枚举
        /// </summary>
        public enum DataViewType
        {
            /// <summary>无数据</summary>
            None,
            /// <summary>识别数据</summary>
            Identification,
            /// <summary>统计数据</summary>
            Statistics,
            /// <summary>涂装下线总数</summary>
            PaintingTotal,
            /// <summary>一次下线不良率/成品率</summary>
            FirstPassYield,
            /// <summary>一检漏检率</summary>
            MissedInspection,
            /// <summary>入库包装数</summary>
            Packaging
        }

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

        private ObservableCollection<PaintingTotalModel> _paintingTotalDatas;
        /// <summary>
        /// 涂装下线总数数据
        /// </summary>
        public ObservableCollection<PaintingTotalModel> PaintingTotalDatas
        {
            get { return _paintingTotalDatas; }
            set { SetProperty(ref _paintingTotalDatas, value); }
        }

        private ObservableCollection<FirstPassYieldModel> _firstPassYieldDatas;
        /// <summary>
        /// 一次下线不良率/成品率数据
        /// </summary>
        public ObservableCollection<FirstPassYieldModel> FirstPassYieldDatas
        {
            get { return _firstPassYieldDatas; }
            set { SetProperty(ref _firstPassYieldDatas, value); }
        }

        private ObservableCollection<MissedInspectionModel> _missedInspectionDatas;
        public ObservableCollection<MissedInspectionModel> MissedInspectionDatas
        {
            get { return _missedInspectionDatas; }
            set { SetProperty(ref _missedInspectionDatas, value); }
        }

        private ObservableCollection<PackagingModel> _packagingDatas;
        public ObservableCollection<PackagingModel> PackagingDatas
        {
            get { return _packagingDatas; }
            set { SetProperty(ref _packagingDatas, value); }
        }

        private DataViewType _currentDataView;
        /// <summary>
        /// 当前显示的数据视图类型（替代多个 Visibility 属性，用 ContentControl+TemplateSelector 切换）
        /// </summary>
        public DataViewType CurrentDataView
        {
            get { return _currentDataView; }
            set
            {
                if (SetProperty(ref _currentDataView, value))
                {
                    HasData = value != DataViewType.None;
                    UpdateQueryCountForView(value);
                }
            }
        }

        private bool _hasData;
        /// <summary>
        /// 是否有数据显示（用于空状态提示）
        /// </summary>
        public bool HasData
        {
            get { return _hasData; }
            set { SetProperty(ref _hasData, value); }
        }

        private void UpdateQueryCountForView(DataViewType viewType)
        {
            switch (viewType)
            {
                case DataViewType.Identification:
                    QueryCount = IdentificationDatas?.Count ?? 0;
                    break;
                case DataViewType.Statistics:
                    QueryCount = StatisticsDatas?.Count ?? 0;
                    break;
                case DataViewType.PaintingTotal:
                    // 涂装总数减去合计行
                    QueryCount = PaintingTotalDatas != null
                        ? PaintingTotalDatas.Count(p => !p.IsTotalRow)
                        : 0;
                    break;
                case DataViewType.FirstPassYield:
                    QueryCount = FirstPassYieldDatas != null
                        ? FirstPassYieldDatas.Count(p => !p.IsTotalRow)
                        : 0;
                    break;
                case DataViewType.MissedInspection:
                    QueryCount = MissedInspectionDatas != null
                        ? MissedInspectionDatas.Count(p => !p.IsTotalRow)
                        : 0;
                    break;
                case DataViewType.Packaging:
                    QueryCount = PackagingDatas != null
                        ? PackagingDatas.Count(p => !p.IsTotalRow)
                        : 0;
                    break;
                default:
                    QueryCount = 0;
                    break;
            }
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

        // 多条件查询筛选属性
        private string _filterModel;
        public string FilterModel
        {
            get { return _filterModel; }
            set { SetProperty(ref _filterModel, value); }
        }

        private string _filterWheelStyle;
        public string FilterWheelStyle
        {
            get { return _filterWheelStyle; }
            set { SetProperty(ref _filterWheelStyle, value); }
        }

        private ObservableCollection<string> _filterStations;
        public ObservableCollection<string> FilterStations
        {
            get { return _filterStations; }
            set { SetProperty(ref _filterStations, value); }
        }

        private string _filterResult;
        public string FilterResult
        {
            get { return _filterResult; }
            set { SetProperty(ref _filterResult, value); }
        }

        private string _filterReportWay;
        public string FilterReportWay
        {
            get { return _filterReportWay; }
            set { SetProperty(ref _filterReportWay, value); }
        }

        private string _filterRemark;
        public string FilterRemark
        {
            get { return _filterRemark; }
            set { SetProperty(ref _filterRemark, value); }
        }

        private int _queryCount;
        public int QueryCount
        {
            get { return _queryCount; }
            set { SetProperty(ref _queryCount, value); }
        }
        private bool _resultBool;
        public bool ResultBool
        {
            get { return _resultBool; }
            set { SetProperty(ref _resultBool, value); }
        }

        // 下拉选项集合
        public ObservableCollection<string> WheelStyles { get; set; }
            = new ObservableCollection<string> { "全部", "成品", "半成品" };
        public ObservableCollection<string> ResultOptions { get; set; }
            = new ObservableCollection<string> { "全部", "合格", "不合格" };
        public ObservableCollection<string> ReportWayOptions { get; set; }
            = new ObservableCollection<string> { "全部", "线上", "平板" };
        public ObservableCollection<string> StationOptions { get; set; }
            = new ObservableCollection<string>();

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
        /// <summary>
        /// 多条件查询命令
        /// </summary>
        public DelegateCommand QueryCommand { get; set; }
        /// <summary>
        /// 重置筛选条件
        /// </summary>
        public DelegateCommand ResetFilterCommand { get; set; }
        /// <summary>
        /// 导出当前查询/统计数据为CSV
        /// </summary>
        public DelegateCommand ExportCurrentDataCommand { get; set; }
        /// <summary>
        /// 查询涂装下线总数
        /// </summary>
        public DelegateCommand QueryPaintingTotalCommand { get; set; }
        /// <summary>
        /// 查询一次下线不良率/成品率
        /// </summary>
        public DelegateCommand QueryFirstPassYieldCommand { get; set; }
        public DelegateCommand QueryMissedInspectionCommand { get; set; }
        public DelegateCommand QueryPackagingCommand { get; set; }
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
            QueryCommand = new DelegateCommand(QueryData);
            ResetFilterCommand = new DelegateCommand(ResetFilter);
            ExportCurrentDataCommand = new DelegateCommand(ExportCurrentData);
            QueryPaintingTotalCommand = new DelegateCommand(QueryPaintingTotal);
            QueryFirstPassYieldCommand = new DelegateCommand(QueryFirstPassYield);
            QueryMissedInspectionCommand = new DelegateCommand(QueryMissedInspection);
            QueryPackagingCommand = new DelegateCommand(QueryPackaging);

            IdentificationDatas = new ObservableCollection<Tbl_productiondatamodel>();
            StatisticsDatas = new ObservableCollection<StatisticsDataModel>();
            PaintingTotalDatas = new ObservableCollection<PaintingTotalModel>();
            FirstPassYieldDatas = new ObservableCollection<FirstPassYieldModel>();
            MissedInspectionDatas = new ObservableCollection<MissedInspectionModel>();
            PackagingDatas = new ObservableCollection<PackagingModel>();

            ProgreVisibility = Visibility.Hidden;
            PickProgreVisibility = Visibility.Hidden;
            Start();

            // 初始化筛选条件默认值
            FilterWheelStyle = "全部";
            FilterStations = new ObservableCollection<string> { "全部" };
            FilterResult = "全部";
            FilterReportWay = "全部";

            // 加载工站列表
            LoadStationOptions();
        }



        /// <summary>
        /// 更新数据
        /// </summary>
        private async void UpdataRecord()
        {
            if (SelectPath != null)
            {

                ProgreVisibility = Visibility.Visible;
                await Task.Run(() =>
                {

                    string[] subDirectories = Directory.GetDirectories(SelectPath);

                    foreach (string subDir in subDirectories)
                    {

                        string WheelPath = FileRenamer.RenameDirectory(subDir);
                        string folderName = Path.GetFileName(WheelPath); //文件夹名称                                   

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
                        int sum = 0;
                        //  处理每个文件
                        foreach (string filePath in files)
                        {
                            // 获取文件名和扩展名
                            string fileName = Path.GetFileName(filePath);
                            bool isTrue = fileName.StartsWith(model) && fileName.Contains(strs[1]); //分类是否准确
                            if (!isTrue)
                            {
                                sum++;
                                //需要调整数据
                                UpdateModelByImageFileName(fileName, filePath, model, style);
                            }
                        }
                        Console.WriteLine($"文件夹：{folderName} 文件总数：{files.Length} 修改数：{sum} ");
                    }
                });

                ProgreVisibility = Visibility.Hidden;
                SelectPath = null;
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
                //Console.WriteLine(result.ToString("yyyy-MM-dd HH:mm:ss"));
                DateTime startTime = result.AddMinutes(-5);
                DateTime endTime = result.AddMinutes(5);


                using (SqlSugarClient db = new SqlAccess().SystemDataAccess)
                {
                    var record = db.Queryable<Tbl_productiondatamodel>()
                                .Where(t => t.RecognitionTime > startTime && t.RecognitionTime < endTime)
                                .Where(t => t.ImagePath != null && t.ImagePath.Contains(fileName))
                                .First();
                    //开启Sql日志输出（调试用）
                    //db.Aop.OnLogExecuting = (sql, pars) =>
                    //{
                    //    Console.WriteLine("---" + sql);
                    //};

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

            CurrentDataView = DataViewType.Identification;
        }
        private void DataInquire()
        {
            if (StartDateTime == null)
                return;
            if (EndDateTime == null)
                return;

            var pDB = new SqlAccess().SystemDataAccess;
            var productionList = pDB.Queryable<Tbl_productiondatamodel>()
                .Where(it => it.RecognitionTime > StartDateTime && it.RecognitionTime <= EndDateTime)
                .OrderBy((sc) => sc.ID, OrderByType.Desc).ToList();
            IdentificationDatas?.Clear();
            IdentificationDatas = new ObservableCollection<Tbl_productiondatamodel>(productionList);
            CurrentDataView = DataViewType.Identification;
            pDB.Close(); pDB.Dispose();
        }

        /// <summary>
        /// 多条件数据查询
        /// </summary>
        private void QueryData()
        {
            if (StartDateTime == null || EndDateTime == null)
            {
                MessageBox.Show("请选择时间范围", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var pDB = new SqlAccess().SystemDataAccess;

            var query = pDB.Queryable<Tbl_productiondatamodel>()
                .Where(it => it.RecognitionTime > StartDateTime && it.RecognitionTime <= EndDateTime);

            // 多条件查询
            if (!string.IsNullOrWhiteSpace(FilterModel))
                query = query.Where(it => it.Model.Contains(FilterModel));

            if (!string.IsNullOrWhiteSpace(FilterWheelStyle) && FilterWheelStyle != "全部")
                query = query.Where(it => it.WheelStyle == FilterWheelStyle);

            if (FilterStations != null && FilterStations.Count > 0 && !FilterStations.Contains("全部"))
            {
                var stations = FilterStations.ToList();
                query = query.Where(it => stations.Contains(it.Station));
            }

            if (!string.IsNullOrWhiteSpace(FilterResult) && FilterResult != "全部")
                query = query.Where(it => it.ResultBool == (FilterResult == "合格"));

            if (!string.IsNullOrWhiteSpace(FilterReportWay) && FilterReportWay != "全部")
                query = query.Where(it => it.ReportWay == FilterReportWay);

            if (!string.IsNullOrWhiteSpace(FilterRemark))
                query = query.Where(it => it.Remark == FilterRemark);

            var productionList = query.OrderBy(it => it.ID, OrderByType.Desc).ToList();

            IdentificationDatas?.Clear();
            IdentificationDatas = new ObservableCollection<Tbl_productiondatamodel>(productionList);
            QueryCount = productionList.Count;
            CurrentDataView = DataViewType.Identification;
            pDB.Close(); pDB.Dispose();
        }

        /// <summary>
        /// 重置筛选条件
        /// </summary>
        private void ResetFilter()
        {
            StartDateTime = DateTime.Now.AddHours(-8);
            EndDateTime = DateTime.Now;
            FilterModel = string.Empty;
            FilterWheelStyle = "全部";
            FilterStations = new ObservableCollection<string> { "全部" };
            FilterResult = "全部";
            FilterReportWay = "全部";
            FilterRemark = string.Empty;
        }

        /// <summary>
        /// 导出当前查询/统计数据为CSV
        /// </summary>
        private void ExportCurrentData()
        {
            // 根据当前视图类型导出对应数据
            switch (CurrentDataView)
            {
                case DataViewType.Identification:
                    if (IdentificationDatas != null && IdentificationDatas.Count > 0)
                        ExportToCsv(IdentificationDatas.ToList());
                    else
                        goto default;
                    break;
                case DataViewType.Statistics:
                    if (StatisticsDatas != null && StatisticsDatas.Count > 0)
                        ExportStatisticsToCsv(StatisticsDatas.ToList());
                    else
                        goto default;
                    break;
                case DataViewType.PaintingTotal:
                    if (PaintingTotalDatas != null && PaintingTotalDatas.Count > 0)
                        ExportPaintingTotalToCsv(PaintingTotalDatas.ToList());
                    else
                        goto default;
                    break;
                case DataViewType.FirstPassYield:
                    if (FirstPassYieldDatas != null && FirstPassYieldDatas.Count > 0)
                        ExportFirstPassYieldToCsv(FirstPassYieldDatas.ToList());
                    else
                        goto default;
                    break;
                case DataViewType.MissedInspection:
                    if (MissedInspectionDatas != null && MissedInspectionDatas.Count > 0)
                        ExportMissedInspectionToCsv(MissedInspectionDatas.ToList());
                    else
                        goto default;
                    break;
                case DataViewType.Packaging:
                    if (PackagingDatas != null && PackagingDatas.Count > 0)
                        ExportPackagingToCsv(PackagingDatas.ToList());
                    else
                        goto default;
                    break;
                default:
                    MessageBox.Show("当前没有可导出的数据，请先查询或统计！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    break;
            }
        }

        /// <summary>
        /// 查询涂装下线总数（精车1号或精车2号的工站数据，按Model+WheelStyle归类）
        /// </summary>
        private void QueryPaintingTotal()
        {
            if (StartDateTime == null || EndDateTime == null)
            {
                MessageBox.Show("请选择时间范围", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var pDB = new SqlAccess().SystemDataAccess;
            var list = pDB.Queryable<Tbl_productiondatamodel>()
                .Where(it => it.RecognitionTime > StartDateTime && it.RecognitionTime <= EndDateTime)
                .Where(it => it.Station == "精车1号" || it.Station == "精车2号")
                .ToList();
            pDB.Close(); pDB.Dispose();

            // 按Model+WheelStyle归类统计
            var grouped = list
                .GroupBy(it => new { Model = it.Model ?? "无型号", WheelStyle = it.WheelStyle ?? "" })
                .OrderBy(g => g.Key.Model)
                .ThenBy(g => g.Key.WheelStyle)
                .ToList();

            PaintingTotalDatas?.Clear();
            int totalCount = 0;
            int index = 1;
            foreach (var group in grouped)
            {
                PaintingTotalDatas.Add(new PaintingTotalModel
                {
                    Index = index.ToString(),
                    Model = group.Key.Model,
                    WheelStyle = group.Key.WheelStyle,
                    WheelCount = group.Count(),
                    IsTotalRow = false
                });
                totalCount += group.Count();
                index++;
            }

            // 添加合计行
            PaintingTotalDatas.Add(new PaintingTotalModel
            {
                Index = "合计",
                Model = string.Empty,
                WheelStyle = string.Empty,
                WheelCount = totalCount,
                IsTotalRow = true
            });

            CurrentDataView = DataViewType.PaintingTotal;
        }

        /// <summary>
        /// 一次下线不良率/成品率
        /// </summary>
        private void QueryFirstPassYield()
        {
            if (StartDateTime == null || EndDateTime == null)
            {
                MessageBox.Show("请选择时间范围", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var pDB = new SqlAccess().SystemDataAccess;
            // 一次性查出所有相关工站的数据
            var allData = pDB.Queryable<Tbl_productiondatamodel>()
                .Where(it => it.RecognitionTime > StartDateTime && it.RecognitionTime <= EndDateTime)
                .Where(it => it.Station == "精车1号" || it.Station == "精车2号"
                    || it.Station == "二检1号" || it.Station == "返修1号")
                .ToList();
            pDB.Close(); pDB.Dispose();

            // 基础分组：精车1号+精车2号，按Model+WheelStyle
            var baseData = allData
                .Where(it => it.Station == "精车1号" || it.Station == "精车2号")
                .GroupBy(it => new { Model = it.Model ?? "无型号", WheelStyle = it.WheelStyle ?? "" })
                .OrderBy(g => g.Key.Model)
                .ThenBy(g => g.Key.WheelStyle)
                .ToList();

            FirstPassYieldDatas?.Clear();
            int totalWheelCount = 0;
            int totalInsBad2 = 0, totalLatheBad1 = 0, totalLatheBad2 = 0, totalRepairOk = 0, totalBad = 0;
            int index = 1;

            foreach (var group in baseData)
            {
                string model = group.Key.Model;
                string style = group.Key.WheelStyle;
                int wheelCount = group.Count();

                // 二检1号不良
                int insBad2 = allData.Count(it =>
                    it.Station == "二检1号" && (it.Model ?? "无型号") == model
                    && it.WheelStyle == style && !it.ResultBool);

                // 精车1号不良
                int latheBad1 = allData.Count(it =>
                    it.Station == "精车1号" && (it.Model ?? "无型号") == model
                    && it.WheelStyle == style && !it.ResultBool);

                // 精车2号不良
                int latheBad2 = allData.Count(it =>
                    it.Station == "精车2号" && (it.Model ?? "无型号") == model
                    && it.WheelStyle == style && !it.ResultBool);

                // 返修合格数
                int repairOk = allData.Count(it =>
                    it.Station == "返修1号" && (it.Model ?? "无型号") == model
                    && it.WheelStyle == style && it.ResultBool);

                int badCount = insBad2 + latheBad1 + latheBad2 - repairOk;

                double badRate = wheelCount > 0 ? (double)badCount / wheelCount : 0;
                // 不良率不能为负数
                if (badRate < 0) badRate = 0;
                double yieldRate = 1 - badRate;

                FirstPassYieldDatas.Add(new FirstPassYieldModel
                {
                    Index = index.ToString(),
                    Model = model,
                    WheelStyle = style,
                    WheelCount = wheelCount,
                    InspectionBad2 = insBad2,
                    LatheBad1 = latheBad1,
                    LatheBad2 = latheBad2,
                    RepairOk = repairOk,
                    BadCount = badCount,
                    CoatingTotal = wheelCount,
                    BadRate = badRate,
                    YieldRate = yieldRate,
                    IsTotalRow = false
                });

                totalWheelCount += wheelCount;
                totalInsBad2 += insBad2;
                totalLatheBad1 += latheBad1;
                totalLatheBad2 += latheBad2;
                totalRepairOk += repairOk;
                totalBad += badCount;
                index++;
            }

            double totalBadRate = totalWheelCount > 0 ? (double)totalBad / totalWheelCount : 0;
            if (totalBadRate < 0) totalBadRate = 0;
            double totalYieldRate = 1 - totalBadRate;

            // 合计行
            FirstPassYieldDatas.Add(new FirstPassYieldModel
            {
                Index = "合计",
                Model = string.Empty,
                WheelStyle = string.Empty,
                WheelCount = totalWheelCount,
                InspectionBad2 = totalInsBad2,
                LatheBad1 = totalLatheBad1,
                LatheBad2 = totalLatheBad2,
                RepairOk = totalRepairOk,
                BadCount = totalBad,
                CoatingTotal = totalWheelCount,
                BadRate = totalBadRate,
                YieldRate = totalYieldRate,
                IsTotalRow = true
            });

            CurrentDataView = DataViewType.FirstPassYield;
        }

        /// <summary>
        /// 一检漏检率
        /// </summary>
        private void QueryMissedInspection()
        {
            if (StartDateTime == null || EndDateTime == null)
            {
                MessageBox.Show("请选择时间范围", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var pDB = new SqlAccess().SystemDataAccess;
            var allData = pDB.Queryable<Tbl_productiondatamodel>()
                .Where(it => it.RecognitionTime > StartDateTime && it.RecognitionTime <= EndDateTime)
                .Where(it => it.Station == "精车1号" || it.Station == "精车2号" || it.Station == "二检1号")
                .ToList();
            pDB.Close(); pDB.Dispose();

            var baseData = allData
                .Where(it => it.Station == "精车1号" || it.Station == "精车2号")
                .GroupBy(it => new { Model = it.Model ?? "无型号", WheelStyle = it.WheelStyle ?? "" })
                .OrderBy(g => g.Key.Model)
                .ThenBy(g => g.Key.WheelStyle)
                .ToList();

            MissedInspectionDatas?.Clear();
            int totalWheelCount = 0, totalInsBad2 = 0;
            int index = 1;

            foreach (var group in baseData)
            {
                string model = group.Key.Model;
                string style = group.Key.WheelStyle;
                int wheelCount = group.Count();

                int insBad2 = allData.Count(it =>
                    it.Station == "二检1号" && (it.Model ?? "无型号") == model
                    && it.WheelStyle == style && !it.ResultBool);

                double missRate = wheelCount > 0 ? (double)insBad2 / wheelCount : 0;

                MissedInspectionDatas.Add(new MissedInspectionModel
                {
                    Index = index.ToString(),
                    Model = model,
                    WheelStyle = style,
                    WheelCount = wheelCount,
                    InspectionBad2 = insBad2,
                    CoatingTotal = wheelCount,
                    MissRate = missRate,
                    IsTotalRow = false
                });

                totalWheelCount += wheelCount;
                totalInsBad2 += insBad2;
                index++;
            }

            double totalMissRate = totalWheelCount > 0 ? (double)totalInsBad2 / totalWheelCount : 0;

            MissedInspectionDatas.Add(new MissedInspectionModel
            {
                Index = "合计",
                Model = string.Empty,
                WheelStyle = string.Empty,
                WheelCount = totalWheelCount,
                InspectionBad2 = totalInsBad2,
                CoatingTotal = totalWheelCount,
                MissRate = totalMissRate,
                IsTotalRow = true
            });

            CurrentDataView = DataViewType.MissedInspection;
        }

        /// <summary>
        /// 入库包装数
        /// </summary>
        private void QueryPackaging()
        {
            if (StartDateTime == null || EndDateTime == null)
            {
                MessageBox.Show("请选择时间范围", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var pDB = new SqlAccess().SystemDataAccess;
            var allData = pDB.Queryable<Tbl_productiondatamodel>()
                .Where(it => it.RecognitionTime > StartDateTime && it.RecognitionTime <= EndDateTime)
                .Where(it => it.Station == "二检1号")
                .ToList();
            pDB.Close(); pDB.Dispose();

            var grouped = allData
                .GroupBy(it => new { Model = it.Model ?? "无型号", WheelStyle = it.WheelStyle ?? "" })
                .OrderBy(g => g.Key.Model)
                .ThenBy(g => g.Key.WheelStyle)
                .ToList();

            PackagingDatas?.Clear();
            int totalInspectionTotal = 0, totalInsBad2 = 0, totalInsOk = 0;
            int index = 1;

            foreach (var group in grouped)
            {
                string model = group.Key.Model;
                string style = group.Key.WheelStyle;
                int inspectionTotal = group.Count();
                int insBad2 = group.Count(it => !it.ResultBool);
                int insOk = inspectionTotal - insBad2;
                double okRate = inspectionTotal > 0 ? (double)insOk / inspectionTotal : 0;

                PackagingDatas.Add(new PackagingModel
                {
                    Index = index.ToString(),
                    Model = model,
                    WheelStyle = style,
                    InspectionTotal = inspectionTotal,
                    InspectionBad2 = insBad2,
                    InspectionOk = insOk,
                    InspectionOkRate = okRate,
                    IsTotalRow = false
                });

                totalInspectionTotal += inspectionTotal;
                totalInsBad2 += insBad2;
                totalInsOk += insOk;
                index++;
            }

            double totalOkRate = totalInspectionTotal > 0 ? (double)totalInsOk / totalInspectionTotal : 0;

            PackagingDatas.Add(new PackagingModel
            {
                Index = "合计",
                Model = string.Empty,
                WheelStyle = string.Empty,
                InspectionTotal = totalInspectionTotal,
                InspectionBad2 = totalInsBad2,
                InspectionOk = totalInsOk,
                InspectionOkRate = totalOkRate,
                IsTotalRow = true
            });

            CurrentDataView = DataViewType.Packaging;
        }

        private void ExportToCsv(List<Tbl_productiondatamodel> dataList)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV文件 (*.csv)|*.csv",
                DefaultExt = ".csv",
                FileName = $"识别数据_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (saveFileDialog.ShowDialog() != true)
                return;

            try
            {
                using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                {
                    // 写入CSV表头
                    writer.WriteLine("序号,轮毂型号,轮毂样式,轮毂高度,轮毂温度,工站,检查结果,NG编号,上报方式,相似度,识别时间");

                    foreach (var item in dataList)
                    {
                        string result = item.ResultBool ? "合格" : "不合格";
                        string remark = item.Remark == "-1" ? "合格" : item.Remark;
                        string time = item.RecognitionTime.ToString("yyyy/MM/dd HH:mm:ss");

                        // CSV转义处理
                        writer.WriteLine($"{item.ID},{EscapeCsvField(item.Model)},{EscapeCsvField(item.WheelStyle)}," +
                            $"{item.WheelHeight},{item.WheelTemperature},{EscapeCsvField(item.Station)}," +
                            $"{result},{EscapeCsvField(remark)},{EscapeCsvField(item.ReportWay)}," +
                            $"{item.Similarity},{time}");
                    }
                }

                MessageBox.Show($"导出成功！共 {dataList.Count} 条记录。\n文件路径：{saveFileDialog.FileName}", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportStatisticsToCsv(List<StatisticsDataModel> dataList)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV文件 (*.csv)|*.csv",
                DefaultExt = ".csv",
                FileName = $"统计数据_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (saveFileDialog.ShowDialog() != true)
                return;

            try
            {
                using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                {
                    // 写入CSV表头
                    writer.WriteLine("序号,轮毂型号,样式,轮毂数量,合格数,主要NG");

                    foreach (var item in dataList)
                    {
                        writer.WriteLine($"{item.Index},{EscapeCsvField(item.Model)},{EscapeCsvField(item.WheelStyle)}," +
                            $"{item.WheelCount},{item.PassCount},{EscapeCsvField(item.MostOfNG)}");
                    }
                }

                MessageBox.Show($"导出成功！共 {dataList.Count} 条记录。\n文件路径：{saveFileDialog.FileName}", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportPaintingTotalToCsv(List<PaintingTotalModel> dataList)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV文件 (*.csv)|*.csv",
                DefaultExt = ".csv",
                FileName = $"涂装下线总数_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (saveFileDialog.ShowDialog() != true)
                return;

            try
            {
                using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                {
                    writer.WriteLine("序号,轮毂型号,样式,轮毂数量");

                    foreach (var item in dataList)
                    {
                        writer.WriteLine($"{EscapeCsvField(item.Index)},{EscapeCsvField(item.Model)}," +
                            $"{EscapeCsvField(item.WheelStyle)},{item.WheelCount}");
                    }
                }

                MessageBox.Show($"导出成功！共 {dataList.Count} 条记录。\n文件路径：{saveFileDialog.FileName}", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportFirstPassYieldToCsv(List<FirstPassYieldModel> dataList)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV文件 (*.csv)|*.csv",
                DefaultExt = ".csv",
                FileName = $"一次下线不良率_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (saveFileDialog.ShowDialog() != true)
                return;

            try
            {
                using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                {
                    writer.WriteLine("序号,轮毂型号,样式,轮毂数量,二检1号不良,精车1号不良,精车2号不良,返修合格数,不良数,涂装总数,不良率,成品率");

                    foreach (var item in dataList)
                    {
                        writer.WriteLine($"{EscapeCsvField(item.Index)},{EscapeCsvField(item.Model)}," +
                            $"{EscapeCsvField(item.WheelStyle)},{item.WheelCount}," +
                            $"{item.InspectionBad2},{item.LatheBad1},{item.LatheBad2}," +
                            $"{item.RepairOk},{item.BadCount},{item.CoatingTotal}," +
                            $"{item.BadRate:P2},{item.YieldRate:P2}");
                    }
                }

                MessageBox.Show($"导出成功！共 {dataList.Count} 条记录。\n文件路径：{saveFileDialog.FileName}", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportMissedInspectionToCsv(List<MissedInspectionModel> dataList)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV文件 (*.csv)|*.csv",
                DefaultExt = ".csv",
                FileName = $"一检漏检率_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (saveFileDialog.ShowDialog() != true)
                return;

            try
            {
                using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                {
                    writer.WriteLine("序号,轮毂型号,样式,轮毂数量,二检1号不良,涂装总数,漏检率");

                    foreach (var item in dataList)
                    {
                        writer.WriteLine($"{EscapeCsvField(item.Index)},{EscapeCsvField(item.Model)}," +
                            $"{EscapeCsvField(item.WheelStyle)},{item.WheelCount}," +
                            $"{item.InspectionBad2},{item.CoatingTotal},{item.MissRate:P2}");
                    }
                }

                MessageBox.Show($"导出成功！共 {dataList.Count} 条记录。\n文件路径：{saveFileDialog.FileName}", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportPackagingToCsv(List<PackagingModel> dataList)
        {
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV文件 (*.csv)|*.csv",
                DefaultExt = ".csv",
                FileName = $"入库包装数_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (saveFileDialog.ShowDialog() != true)
                return;

            try
            {
                using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                {
                    writer.WriteLine("序号,轮毂型号,样式,二检总数,二检1号不良,二检合格数,二检合格率");

                    foreach (var item in dataList)
                    {
                        writer.WriteLine($"{EscapeCsvField(item.Index)},{EscapeCsvField(item.Model)}," +
                            $"{EscapeCsvField(item.WheelStyle)},{item.InspectionTotal}," +
                            $"{item.InspectionBad2},{item.InspectionOk},{item.InspectionOkRate:P2}");
                    }
                }

                MessageBox.Show($"导出成功！共 {dataList.Count} 条记录。\n文件路径：{saveFileDialog.FileName}", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// CSV字段转义（处理逗号、引号、换行）
        /// </summary>
        private string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return string.Empty;

            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

        /// <summary>
        /// 加载工站列表
        /// </summary>
        private void LoadStationOptions()
        {
            try
            {
                var pDB = new SqlAccess().SystemDataAccess;
                var stations = pDB.Queryable<Tbl_productiondatamodel>()
                    .Where(it => !string.IsNullOrEmpty(it.Station))
                    .Select(it => it.Station)
                    .ToList()
                    .Distinct()
                    .OrderBy(s => s)
                    .ToList();

                StationOptions?.Clear();
                StationOptions.Add("全部");
                foreach (var station in stations)
                    StationOptions.Add(station);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载工站列表失败: {ex.Message}");
            }
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

            var pDB = new SqlAccess().SystemDataAccess;
            // 从数据库读取的数据
            var productionList = pDB.Queryable<Tbl_productiondatamodel>()
                                                        .Where(it => it.RecognitionTime > StartDateTime && it.RecognitionTime <= EndDateTime)
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
            CurrentDataView = DataViewType.Statistics;
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
            ProgressDialogController controller = await this._dialogCoordinator.ShowProgressAsync(this, "数据导出", "数据导出到本地中");

            try
            {
                if (StartDateTime == null || EndDateTime == null)
                {
                    await _dialogCoordinator.ShowMessageAsync(this, "提示", "请先设置开始时间和结束时间！");
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
                controller.SetIndeterminate();

                //await Task.Delay(3000);       
                await Task.Run(() =>
                {
                    var db = new SqlAccess().SystemDataAccess;
                    var query = db.Queryable<Tbl_productiondatamodel>()
                        .Where(it => it.RecognitionTime > StartDateTime && it.RecognitionTime <= EndDateTime);

                    // 多条件查询
                    if (!string.IsNullOrWhiteSpace(FilterModel))
                        query = query.Where(it => it.Model.Contains(FilterModel));

                    if (!string.IsNullOrWhiteSpace(FilterWheelStyle) && FilterWheelStyle != "全部")
                        query = query.Where(it => it.WheelStyle == FilterWheelStyle);

                    if (FilterStations != null && FilterStations.Count > 0 && !FilterStations.Contains("全部"))
                    {
                        var stations = FilterStations.ToList();
                        query = query.Where(it => stations.Contains(it.Station));
                    }

                    if (!string.IsNullOrWhiteSpace(FilterResult) && FilterResult != "全部")
                        query = query.Where(it => it.ResultBool == (FilterResult == "合格"));

                    if (!string.IsNullOrWhiteSpace(FilterReportWay) && FilterReportWay != "全部")
                        query = query.Where(it => it.ReportWay == FilterReportWay);

                    if (!string.IsNullOrWhiteSpace(FilterRemark))
                        query = query.Where(it => it.Remark == FilterRemark);

                    List<Tbl_productiondatamodel> list = query.ToList();
                    db.Close(); db.Dispose();

                    if (list.Count == 0)
                    {
                        path = string.Empty;
                        return;
                    }

                    ExportProducts("半成品", list, workShift);
                    path = ExportProducts("成品", list, workShift);
                });
               
                if (string.IsNullOrEmpty(path))
                {
                    await _dialogCoordinator.ShowMessageAsync(this, "提示", $"该时间段内没有数据！{StartDateTime} >> {EndDateTime}");
                    return;
                }

                Console.WriteLine($"数据导出完成");
                //EventMessage.SystemMessageDisplay("数据导出完成", MessageType.Success);
                await Task.Delay(100);
                Process.Start("explorer.exe", path);

            }
            catch (Exception ex)
            {
                await _dialogCoordinator.ShowMessageAsync(this, "错误", $"数据导出失败: {ex.Message}");
            }
            finally
            {
                await controller.CloseAsync();
            }
        }


        private async void DataExportExcel()
        {
            ProgressDialogController controller = await this._dialogCoordinator.ShowProgressAsync(this, "数据导出", "数据导出到本地中");

            try
            {
                controller.SetIndeterminate();
                string path = await AsyncExport();

                if (string.IsNullOrEmpty(path))
                {
                    await _dialogCoordinator.ShowMessageAsync(this, "提示", "该时间段内没有数据！");
                    return;
                }

                //Console.WriteLine($"数据导出完成");
                await Task.Delay(100);
                Process.Start("explorer.exe", path);
            }
            catch (Exception ex)
            {
                await _dialogCoordinator.ShowMessageAsync(this, "错误", $"数据导出失败: {ex.Message}");
            }
            finally
            {
                await controller.CloseAsync();
            }
        }

        public async Task<string> AsyncExport()
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

            string resultPath = string.Empty;
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
                resultPath = ExportProducts("成品", list, workShift);
            });

            return resultPath;
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
