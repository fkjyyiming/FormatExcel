using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FormatExcel
{
    public class FileComparer
    {
        public static (int totalCompared, int totalDiff, List<string> diffPdf, List<string> diffDwg) CompareFiles(string pdfFolderPath, string dwgFolderPath)
        {
            // 获取 PDF 文件列表，去除后缀
            List<string> pdfFiles = Directory.GetFiles(pdfFolderPath, "*.pdf", SearchOption.TopDirectoryOnly)
                                            .Select(file => Path.GetFileNameWithoutExtension(file)).ToList();

            // 获取 DWG 文件列表，去除后缀，并排除其他格式的文件
            List<string> dwgFiles = Directory.GetFiles(dwgFolderPath, "*.dwg*", SearchOption.TopDirectoryOnly)
                                                .Where(file => Path.GetExtension(file).Equals(".dwg.dwg", StringComparison.OrdinalIgnoreCase) || Path.GetExtension(file).Equals(".dwg", StringComparison.OrdinalIgnoreCase))
                                                 .Select(file => Path.GetFileNameWithoutExtension(file).Replace(".dwg", "").Replace(".dwg", "")).ToList();


            // 比较文件名
            var diffPdf = pdfFiles.Except(dwgFiles);
            var diffDwg = dwgFiles.Except(pdfFiles);


            // 计算对比数量
            int totalCompared = Math.Min(pdfFiles.Count, dwgFiles.Count);
            int totalDiff = diffPdf.Count() + diffDwg.Count();


            return (totalCompared, totalDiff, diffPdf.ToList(), diffDwg.ToList());
        }
    }
}