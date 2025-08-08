using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQLManage.Util
{
    internal class FileRenamer
    {
        /// <summary>
        /// 文件夹重命名
        /// </summary>
        /// <param name="originalPath"></param>
        /// <param name="newName"></param>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="IOException"></exception>
        public static string RenameDirectory(string originalPath)
        {
            // 获取父目录路径
            string parentDirectory = Path.GetDirectoryName(originalPath);
            string folderName = Path.GetFileName(originalPath); //文件夹名称
            if (folderName.IndexOf('-') != -1)
            {
                folderName = folderName.Replace('-', '_');
            }
            else
            {
                return originalPath;
            }

            if (string.IsNullOrEmpty(parentDirectory))
            {
                throw new ArgumentException("无法从路径中提取父目录");
            }

            // 构建新路径
            string newPath = Path.Combine(parentDirectory, folderName);

            // 检查原始目录是否存在
            if (!Directory.Exists(originalPath))
            {
                throw new DirectoryNotFoundException($"原始目录不存在: {originalPath}");
            }

            // 检查新目录是否已存在
            if (Directory.Exists(newPath))
            {
                throw new IOException($"目标目录已存在: {newPath}");
            }

            // 执行重命名
            Directory.Move(originalPath, newPath);

            return newPath;
        }

        /// <summary>
        /// 文件重命名
        /// </summary>
        /// <param name="rootDirectory"></param>
        public static void RenameFilesInSubdirectories(string rootDirectory)
        {
            try
            {
                // 1. 检查根目录是否存在
                if (!Directory.Exists(rootDirectory))
                {
                    Console.WriteLine($"目录不存在: {rootDirectory}");
                    return;
                }

                // 2. 获取所有子目录
                var subDirectories = Directory.GetDirectories(rootDirectory);
                Console.WriteLine($"在 {rootDirectory} 中找到 {subDirectories.Length} 个子目录");

                int totalRenamed = 0;
                int skipped = 0;

                foreach (var subDir in subDirectories)
                {
                    // 3. 获取子目录名称（不带路径）
                    string folderName = Path.GetFileName(subDir);
                    Console.WriteLine($"\n处理子目录: {folderName}");

                    // 4. 获取目录下所有文件
                    var files = Directory.GetFiles(subDir);
                    if (files.Length == 0)
                    {
                        Console.WriteLine("  -- 没有找到文件");
                        continue;
                    }

                    // 5. 处理每个文件
                    foreach (var filePath in files)
                    {
                        try
                        {
                            // 6. 获取文件名和扩展名
                            string fileName = Path.GetFileName(filePath);
                            string extension = Path.GetExtension(filePath);

                            // 7. 提取原始文件名中的日期部分
                            string datePart = ExtractDatePart(fileName);

                            if (string.IsNullOrEmpty(datePart))
                            {
                                Console.WriteLine($"  -- 跳过: {fileName} (未找到日期部分)");
                                skipped++;
                                continue;
                            }

                            // 8. 构建新文件名
                            string newFileName = $"{folderName}&{datePart}{extension}";
                            string newFilePath = Path.Combine(subDir, newFileName);

                            // 9. 重命名文件
                            if (File.Exists(newFilePath))
                            {
                                Console.WriteLine($"  -- 跳过: {newFileName} (目标文件已存在)");
                                skipped++;
                            }
                            else
                            {
                                File.Move(filePath, newFilePath);
                                Console.WriteLine($"  ++ 重命名: {fileName} -> {newFileName}");
                                totalRenamed++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  !! 处理文件 {filePath} 时出错: {ex.Message}");
                        }
                    }
                }

                // 10. 输出统计信息
                Console.WriteLine($"\n操作完成! 重命名: {totalRenamed} 个文件, 跳过: {skipped} 个文件");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理过程中发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 从文件名中提取日期部分 (YYYYMMMDD)
        /// </summary>
        private static string ExtractDatePart(string fileName)
        {
            // 查找最后一个 '&' 符号的位置
            int lastAmpersandIndex = fileName.LastIndexOf('&');

            if (lastAmpersandIndex < 0 || lastAmpersandIndex >= fileName.Length - 1)
            {
                return null; // 没有找到 & 符号
            }

            // 获取 & 符号后的部分（不含扩展名）
            string datePart = fileName.Substring(lastAmpersandIndex + 1);

            // 移除扩展名（如果存在）
            int dotIndex = datePart.LastIndexOf('.');
            if (dotIndex > 0)
            {
                datePart = datePart.Substring(0, dotIndex);
            }

            return datePart;
        }

        // 使用示例
        //public static void Main()
        //{
        //    string directoryPath = @"D:\ZS\终检";
        //    RenameFilesInSubdirectories(directoryPath);

        //    // 按任意键退出
        //    Console.WriteLine("\n按任意键退出...");
        //    Console.ReadKey();
        //}
    }
}
