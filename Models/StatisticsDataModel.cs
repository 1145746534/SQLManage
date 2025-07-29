using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQLManage.Models
{
    internal class StatisticsDataModel
    {
        public int Index { get; set; }
        /// <summary>
        /// 轮型
        /// </summary>
        public string Model { get; set; }
        /// <summary>
        /// 轮毂数量
        /// </summary>
        public int WheelCount { get; set; }

        /// <summary>
        /// 轮毂样式（成品/半成品）
        /// </summary>
        public string WheelStyle { get; set; }

        /// <summary>
        /// 合格数
        /// </summary>
        public int PassCount { get; set; }

        /// <summary>
        /// NG（排名前3）
        /// </summary>
        public string MostOfNG { get; set; }
    }
}
