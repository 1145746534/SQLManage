namespace SQLManage.Models
{
    internal class FirstPassYieldModel
    {
        public string Index { get; set; }
        public string Model { get; set; }
        public string WheelStyle { get; set; }
        /// <summary>轮毂数量（精车1号+精车2号）</summary>
        public int WheelCount { get; set; }
        /// <summary>二检1号不良</summary>
        public int InspectionBad2 { get; set; }
        /// <summary>精车1号不良</summary>
        public int LatheBad1 { get; set; }
        /// <summary>精车2号不良</summary>
        public int LatheBad2 { get; set; }
        /// <summary>返修合格数</summary>
        public int RepairOk { get; set; }
        /// <summary>不良数 = 二检不良+精车1不良+精车2不良-返修合格数</summary>
        public int BadCount { get; set; }
        /// <summary>涂装总数（精车1号+精车2号）</summary>
        public int CoatingTotal { get; set; }
        /// <summary>不良率</summary>
        public double BadRate { get; set; }
        /// <summary>成品率 = 1-不良率</summary>
        public double YieldRate { get; set; }
        public bool IsTotalRow { get; set; }
    }
}
