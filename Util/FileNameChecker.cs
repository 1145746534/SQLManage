using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQLManage.Util
{
    internal static class FileNameChecker
    {
        /// <summary>
        /// 挑选不同于上一级目录名称的文件
        /// </summary>
        /// <param name="sourceRoot"></param>
        /// <param name="targetRoot"></param>
        public async static Task<bool> CheckAndExportFolderStructure(string sourceRoot, string targetRoot)
        {
            await Task.Run(() =>
            {


                var subDirectories = Directory.GetDirectories(sourceRoot);

                foreach (string subDir in subDirectories)
                {
                    string WheelPath = FileRenamer.RenameDirectory(subDir);

                    string folderName = Path.GetFileName(WheelPath); //文件夹名称
                    int separatorIndex = folderName.IndexOf('_');
                    if (separatorIndex == -1)
                    {
                        Console.WriteLine($"跳过格式错误的文件夹: {folderName}");

                    }
                    string[] files = Directory.GetFiles(WheelPath);
                    string filePath = string.Empty;
                    foreach (var file in files)
                    {
                        string fileName = Path.GetFileName(file);  //文件
                        if (!CheckFileNameConsistency(fileName, folderName))
                        {
                            //文件复制转移
                            ExportFile(file, targetRoot, folderName);

                            //hasInconsistency = true;
                            //filePath

                        }
                    }


                }

            });
            return true;
        }

        private static string ExtractWheelFromFolderName(string folderName)
        {
            int separatorIndex = folderName.IndexOf('_');
            if (separatorIndex == -1)
            {
                separatorIndex = folderName.IndexOf('-');
            }
            return separatorIndex > 0 ? folderName.Substring(0, separatorIndex) : null;
        }

        private static bool CheckFileNameConsistency(string fileName, string folderWheel)
        {
            int separatorIndex = fileName.IndexOf('&');
            string fileWheel = separatorIndex > 0 ? fileName.Substring(0, separatorIndex) : null;

            return fileWheel.Contains(folderWheel);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sourcefile"></param>
        /// <param name="targetRoot"></param>
        /// <param name="folderName"></param>
        private static void ExportFile(string sourcefile, string targetRoot, string folderName)
        {
            folderName = folderName.Replace('-', '_');
            string targetDir = Path.Combine(targetRoot, folderName);

            Directory.CreateDirectory(targetDir);



            string targetFile = Path.Combine(targetDir, Path.GetFileName(sourcefile));
            File.Copy(sourcefile, targetFile, overwrite: true);


            //Console.WriteLine($"已导出不一致文件夹: {relativePath}");
        }

        // 添加 .NET Framework 兼容的相对路径计算方法
        private static string MakeRelativePath(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath)) return toPath;
            if (string.IsNullOrEmpty(toPath)) return string.Empty;

            if (fromPath[fromPath.Length - 1] != Path.DirectorySeparatorChar)
                fromPath += Path.DirectorySeparatorChar;

            Uri fromUri = new Uri(fromPath);
            Uri toUri = new Uri(toPath);

            if (fromUri.Scheme != toUri.Scheme)
                return toPath;

            Uri relativeUri = fromUri.MakeRelativeUri(toUri);
            string relativePath = Uri.UnescapeDataString(relativeUri.ToString());

            return relativePath.Replace('/', Path.DirectorySeparatorChar);
        }
    }

}
