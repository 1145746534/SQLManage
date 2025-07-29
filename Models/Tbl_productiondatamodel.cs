using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SqlSugar;


namespace SQLManage.Models
{
    /// <summary>
    /// 生产数据模型
    /// </summary>
    internal class Tbl_productiondatamodel
    {
        [SugarColumn(IsPrimaryKey = true, IsIdentity = true)] //IsIdentity为True时Id为自增型
        public int ID { get; set; }

        [SugarColumn(IsNullable = true, Length = 36)]
        public string GUID { get; set; }

        /// <summary>
        /// 匹配轮型名称
        /// </summary>
        [SugarColumn(IsNullable = true, Length = 50)]
        public string WheelType { get; set; }
        /// <summary>
        /// 识别耗时
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true, Length = 20)]
        public string TimeConsumed { get; set; }
        /// <summary>
        /// 相似度
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true, Length = 10)]
        public string Similarity { get; set; }

        /// <summary>
        /// 轮毂高度
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true)]
        public float WheelHeight { get; set; }

        /// <summary>
        /// 轮毂温度
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true)]
        public float WheelTemperature { get; set; }

        /// <summary>
        /// 样式
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true, Length = 10)]
        public string WheelStyle { get; set; }

        /// <summary>
        /// 颜色
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true, Length = 10)]
        public string WheelColor { get; set; }

        /// <summary>
        /// 记录时间
        /// </summary>
        [SugarColumn(IsNullable = false, ColumnDataType = "datetime")]
        public DateTime RecognitionTime { get; set; }
        /// <summary>
        /// 图片路径
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true, Length = 150)]
        public string ImagePath { get; set; }

        /// <summary>
        /// 传输编码
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true, Length = 20)]
        public string TransmissionCoding { get; set; }

        /// <summary>
        /// 轮毂型号
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = false, Length = 32)]
        public string Model { get; set; }
        /// <summary>
        /// 工站
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true, Length = 20)]
        public string Station { get; set; }
        /// <summary>
        /// 下一站（回流/下转/不合格）
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true, Length = 20)]
        public string NextStation { get; set; }

        /// <summary>
        /// 检查结果
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true)]
        public bool ResultBool { get; set; }

        /// <summary>
        /// 检测说明(缺陷编码)
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = false, Length = 5)]
        public string Remark { get; set; }

        /// <summary>
        /// 上报人/报废区使用
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true, Length = 20)]
        public string ReportPerson { get; set; }

        /// <summary>
        /// 上报方式/线上or平板
        /// </summary>
        [SqlSugar.SugarColumn(IsNullable = true, Length = 8)]
        public string ReportWay { get; set; }

        /// <summary>
        /// 上报日期 存储格式：YYYY-MM-DD
        /// </summary>
        [SugarColumn(ColumnDataType = "date")]
        public DateTime RecognitionDay { get; set; }
    }
}
