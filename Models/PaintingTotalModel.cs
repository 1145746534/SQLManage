using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQLManage.Models
{
    internal class PaintingTotalModel
    {
        /// <summary>
        /// 序号（最后一行显示"合计"）
        /// </summary>
        public string Index { get; set; }
        /// <summary>
        /// 轮毂型号
        /// </summary>
        public string Model { get; set; }
        /// <summary>
        /// 样式
        /// </summary>
        public string WheelStyle { get; set; }
        /// <summary>
        /// 轮毂数量
        /// </summary>
        public int WheelCount { get; set; }
        /// <summary>
        /// 是否为合计行
        /// </summary>
        public bool IsTotalRow { get; set; }
    }
}
