namespace SQLManage.Models
{
    internal class PackagingModel
    {
        public string Index { get; set; }
        public string Model { get; set; }
        public string WheelStyle { get; set; }
        /// <summary>二检总数</summary>
        public int InspectionTotal { get; set; }
        /// <summary>二检1号不良</summary>
        public int InspectionBad2 { get; set; }
        /// <summary>二检合格数 = 二检总数 - 二检1号不良</summary>
        public int InspectionOk { get; set; }
        /// <summary>二检合格率 = 二检合格数 / 二检总数</summary>
        public double InspectionOkRate { get; set; }
        public bool IsTotalRow { get; set; }
    }
}
