using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQLManage.Models
{
    internal class GetDateTimeModel
    {
        /// <summary>
        /// 起始时间
        /// </summary>
        public DateTime StartDateTime { get; set; }
        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime EndDateTime { get; set; }
        /// <summary>
        /// 获取结果
        /// </summary>
        public string Result { get; set; }
    }
}
