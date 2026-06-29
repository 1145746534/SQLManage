namespace SQLManage.Models
{
    internal class MissedInspectionModel
    {
        public string Index { get; set; }
        public string Model { get; set; }
        public string WheelStyle { get; set; }
        /// <summary>轮毂数量（精车1号+精车2号）</summary>
        public int WheelCount { get; set; }
        /// <summary>二检1号不良</summary>
        public int InspectionBad2 { get; set; }
        /// <summary>涂装总数</summary>
        public int CoatingTotal { get; set; }
        /// <summary>漏检率 = 二检1号不良 / 涂装总数</summary>
        public double MissRate { get; set; }
        public bool IsTotalRow { get; set; }
    }
}
