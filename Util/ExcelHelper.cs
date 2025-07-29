using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SQLManage.Models;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace SQLManage.Util
{
    internal class ExcelHelper
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="exportDatas"></param>
        /// <param name="sourceFile"></param>
        /// <param name="targetDir"></param>
        public void ModifyExcelFile(Queue<ExportDataModel> exportDatas, string sourceFile, string targetDir)
        {

            Application excelApp = null;
            Workbook workbook = null;

            try
            {
                string sheetName = "模板";
                string path = CopyFileWithDateName(sourceFile, targetDir);

                // 1. 创建Excel应用实例
                excelApp = new Application();
                excelApp.Visible = false; // 后台运行
                excelApp.DisplayAlerts = false; // 禁止弹出警告

                // 2. 打开工作簿（不破坏格式）
                workbook = excelApp.Workbooks.Open(path);

                // 3. 获取指定工作表
                Worksheet worksheet = workbook.Sheets[sheetName] as Worksheet;
                if (worksheet == null)
                {
                    throw new Exception($"工作表 '{sheetName}' 不存在");
                }

                while (exportDatas.Count > 0)
                {
                    ExportDataModel data = exportDatas.Dequeue();
                    //匹配项
                    int matchR = data.MatchRow;
                    int startCol = data.MatchStartCol;
                    int endCol = data.MatchEndCol;
                    string matchN = data.MatchName;
                    //插入项
                    int setR = data.SettingRow;
                    object setV = data.SettingValue;
                    //Console.WriteLine($"{startCol} -> {endCol} - {setR} - {setV}");

                    //循环这个行下所有的列
                    for (int i = startCol; i <= endCol; i++)
                    {
                        // 单元格值（转换为字符串）
                        Excel.Range cell1 = (Excel.Range)worksheet.Cells[matchR, i];
                        string value1 = cell1.Text as string ?? cell1.Value2?.ToString();
                        if (value1 != null && value1.Trim().Equals(matchN, StringComparison.OrdinalIgnoreCase))
                        {
                            //  修改B列值
                            Excel.Range cellB = (Excel.Range)worksheet.Cells[setR, i];
                            cellB.Value2 = setV;  // 只修改值，不改变格式
                            break;
                        }

                    }




                    //ProcessData(data);
                }


                // 4. 获取使用的行范围
                //Excel.Range usedRange = worksheet.UsedRange;
                //int rowCount = usedRange.Rows.Count;

                //// 5. 遍历查找目标行
                //bool found = false;
                //for (int row = 1; row <= rowCount; row++)
                //{
                //    // 获取A列单元格值（转换为字符串）
                //    Excel.Range cellA = (Excel.Range)worksheet.Cells[row, 1];
                //    string valueA = cellA.Text as string ?? cellA.Value2?.ToString();

                //    // 检查是否匹配
                //    if (valueA != null && valueA.Trim().Equals("098q4", StringComparison.OrdinalIgnoreCase))
                //    {
                //        // 6. 修改B列值
                //        Excel.Range cellB = (Excel.Range)worksheet.Cells[row, 2];
                //        cellB.Value2 = "TY";  // 只修改值，不改变格式

                //        Console.WriteLine($"已修改行 {row} 的B列值");
                //        found = true;
                //        break; // 找到第一个匹配项后停止
                //    }

                //}

                //if (!found)
                //{
                //    Console.WriteLine("未找到匹配'A列=098q4'的行");
                //}

                // 7. 保存并关闭（保留原始格式）
                workbook.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"操作失败: {ex.Message}");
            }
            finally
            {
                // 8. 清理资源
                if (workbook != null)
                {
                    workbook.Close(false); // 不保存更改（因为前面已经保存过）
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }

                // 强制垃圾回收释放COM对象
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


        public string CopyFileWithDateName(string sourceFile, string targetDir)
        {

            // 1. 校验源文件是否存在
            if (!File.Exists(sourceFile))
                throw new FileNotFoundException("源文件不存在: " + sourceFile);

            // 2. 创建目标文件夹（若不存在）
            Directory.CreateDirectory(targetDir);

            // 3. 生成新文件名（格式：年月日）
            string dateStr = DateTime.Now.ToString("yyyyMMdd HH-mm-ss"); // 格式示例：20250618[3,7](@ref)
            string originalName = Path.GetFileNameWithoutExtension(sourceFile);
            string extension = Path.GetExtension(sourceFile);
            string newFileName = $"{dateStr}{extension}"; // 保留原名+日期后缀[3](@ref)
            string destPath = Path.Combine(targetDir, newFileName);

            // 4. 执行复制（覆盖同名文件）
            File.Copy(sourceFile, destPath, true); // true 表示覆盖[9,10](@ref)
            Console.WriteLine($"文件已复制并重命名：{destPath}");
            return destPath;


        }
    }
}
