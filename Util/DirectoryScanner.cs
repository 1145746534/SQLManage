using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQLManage.Util
{

    internal class DirectoryScanner
    {
        public static Dictionary<string, string> ScanDirectories(string rootPath)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 检查根目录是否存在
            if (!Directory.Exists(rootPath))
            {
                throw new DirectoryNotFoundException($"根目录不存在: {rootPath}");
            }

            try
            {
                // 遍历一级子目录
                foreach (string firstLevelDir in Directory.GetDirectories(rootPath))
                {
                    string firstLevelName = Path.GetFileName(firstLevelDir);
                    string key = $"{firstLevelName}";
                    result[key] = firstLevelDir;

                    //// 遍历二级子目录
                    //foreach (string secondLevelDir in Directory.GetDirectories(firstLevelDir))
                    //{
                    //    string secondLevelName = Path.GetFileName(secondLevelDir);
                    //    string key = $"{firstLevelName}{secondLevelName}";

                    //    // 添加到字典（自动处理重复键名）
                    //    result[key] = secondLevelDir;
                    //}
                }
                return result;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"目录扫描失败: {ex.Message}", ex);
            }
        }
    }
}
