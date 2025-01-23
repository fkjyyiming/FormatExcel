using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Crypto.Digests;
using SixLabors.Fonts.Tables.AdvancedTypographic;
using SixLabors.ImageSharp.PixelFormats;
using System;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text.RegularExpressions;

namespace FormatExcel
{
    public static class ExcelGenerator
    {
        public static void GenerateExcelReport(string pdfFolderPath, string designStage, string category, string toCompany, string zones, string templatePath, string savePath)
        {
            // 获取 PDF 文件列表
            string[] pdfFiles = Directory.GetFiles(pdfFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);
            IWorkbook workbook = null;
            try
            {
                using (FileStream templateFile = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                {
                    if (templatePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = new HSSFWorkbook(templateFile);
                    }
                    else if (templatePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = new XSSFWorkbook(templateFile);
                    }
                    else
                    {
                        throw new Exception("不支持的模板文件类型!");
                    }
                    if (workbook == null)
                    {
                        throw new Exception("Excel 模板文件打开失败，请检查模板文件");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Excel 模板文件打开失败，请检查模板文件：{ex.Message}");
            }
            using (workbook)
            {
                ISheet worksheet = workbook.GetSheetAt(0); // 获取第一个工作表
                int row = 1; // 从第2行开始填充数据, NPOI 的索引是0，但是我们希望从第二行开始，所以初始值是 1.
                //遍历pdf文件
                foreach (string pdfFile in pdfFiles)
                {
                    try
                    {
                        string pdfFileName = Path.GetFileName(pdfFile);
                        string fileName = Path.GetFileNameWithoutExtension(pdfFileName);

                        string documentNumber = GetDocumentNumber(fileName);
                        string description = GetDescription(fileName);
                        string revisionNumber = GetRevisionNumber(fileName);
                        string discipline = GetDiscipline(fileName);
                        string levels = GetLevels(fileName);
                        string zones_1 = GetZones(fileName);


                        IRow dataRow = worksheet.CreateRow(row);
                        dataRow.CreateCell(0).SetCellValue(documentNumber);
                        dataRow.CreateCell(1).SetCellValue(description);
                        dataRow.CreateCell(2).SetCellValue(revisionNumber);
                        dataRow.CreateCell(3).SetCellValue(designStage);
                        dataRow.CreateCell(4).SetCellValue(discipline);
                        dataRow.CreateCell(5).SetCellValue(category);
                        dataRow.CreateCell(6).SetCellValue(toCompany);
                        if (zones == "")
                        {
                            dataRow.CreateCell(7).SetCellValue(zones_1);
                        }
                        else
                        {
                            dataRow.CreateCell(7).SetCellValue(zones);
                        }
                        
                        dataRow.CreateCell(8).SetCellValue(levels);
                        dataRow.CreateCell(9).SetCellValue($"{pdfFileName},{fileName}.dwg");

                        row++; //递增row
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"文件数据写入异常，请检查命名规则:{ex.Message}");
                    }

                }

                try
                {
                    using (FileStream file = new FileStream(savePath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(file);
                        if (!File.Exists(savePath))
                        {
                            throw new Exception($"文件保存失败，请检查权限，路径是否正确:{savePath}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"保存文件时发生错误，请检查路径是否有效，是否具有写入权限:{ex.Message}");
                }
            }
        }

        private static string GetZones(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return null;

            // 1. 获取文件名，去除后缀
            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(fileName);

            // 2. 使用下划线 "_" 分割文件名
            string[] parts = fileNameWithoutExtension.Split('_');

            // 3. 返回最后一个部分
            if (parts.Length > 0)
            {
                //取名称的最后一部分
                string part = parts.Last() ;
                //将最后一部分用-进行分隔
                string[] lastPart = fileNameWithoutExtension.Split('-');
                //再取其第一部分，如VL1 TV1...
                string unit = parts.First();

                switch (unit)
                {
                    case  "C10":
                    case  "TC10":
                    case  "C10M":
                        return "Villa Type C10";
                    case "VL1":
                    case "TV1":
                    case "V1M":
                        return "Villa Type 01";
                    case "VL2":
                    case "TV2":
                    case "V2M":
                        return "Villa Type 02";
                    case "VL3":
                    case "TV3":
                    case "V3M":
                        return "Villa Type 03";
                    case "VL4":
                    case "TV4":
                    case "V4M":
                        return "Villa Type 04";
                    case "DP1":
                    case "DP1M":
                    case "TDP1":
                        return "Duplex Type 01";
                    case "DP2":
                    case "DP2M":
                    case "TDP2":
                        return "Duplex Type 02";
                    case "DP3":
                    case "DP3M":
                    case "TDP3":
                        return "Duplex Type 03";
                    case "DP4":
                    case "DP4M":
                    case "TDP4":
                        return "Duplex Type 04";
                    case "TH1":
                    case "TH1M":
                    case "TTH1":
                        return "Town House 01";
                    case "TH2":
                    case "TH2M":
                    case "TTH2":
                        return "Town House 02";
                    case "TH3":
                    case "TH3M":
                    case "TTH3":
                        return "Town House 03";
                    case "D01":
                        return "Cluster Type1(DP1-DP1)";
                    case "D02":
                        return "Cluster Type2(DP2-DP2)";
                    case "D04":
                        return "Cluster Type3(DP4-DP4)";
                    case "Q01":
                        return "Cluster Type4(DP1-TH1-TH1-DP1)";
                    case "Q02":
                        return "Cluster Type5(DP1-TH3-TH3-DP1)";
                    case "Q03":
                        return "Cluster Type6(DP2-TH1-TH1-DP2)";
                    case "Q04":
                        return "Cluster Type7(DP2-TH3-TH3-DP2)";
                    case "Q05":
                        return "Cluster Type8(DP4-TH1-TH1-DP4)";
                    case "Q06":
                        return "Cluster Type9(DP4-TH3-TH3-DP4)";
                    case "HX1":
                        return "Cluster Type10(DP1-TH1-TH3-TH3-TH1-DP1)";
                    case "HX2":
                        return "Cluster Type11(DP1-TH3-TH3-TH3-TH3-DP1)";
                    case "HX3":
                        return "Cluster Type12(DP2-TH1-TH1-TH1-TH1-DP2)";
                    case "HX4":
                        return "Cluster Type13(DP2-TH1-TH3-TH3-TH1-DP2)";
                    case "HX5":
                        return "Cluster Type14(DP4-TH1-TH1-TH1-TH1-DP4)";
                    case "HX6":
                        return "Cluster Type15(DP4-TH1-TH3-TH3-TH1-DP4)";
                    case "HX7":
                        return "Cluster Type16(DP4-TH3-TH3-TH3-TH3-DP4)";

                    default:
                        return unit;
                }
            }
            else
            {
                return null; // 如果分割后没有部分，返回null
            }
        }

        // 解析DocumentNumber
        private static string GetDocumentNumber(string fileName)
        {
            string pattern = @"^([^_]+(?:_[^_]+){6})";
            System.Text.RegularExpressions.Match match = Regex.Match(fileName, pattern);
            if (match.Success)
            {
                return match.Groups[1].Value.Replace('_', '-');
            }
            return fileName;
        }

        // 解析Description
        private static string GetDescription(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return null;

            // 1. 获取文件名，去除后缀
            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(fileName);

            // 2. 使用下划线 "_" 分割文件名
            string[] parts = fileNameWithoutExtension.Split('_');

            // 3. 返回最后一个部分
            if (parts.Length > 0)
            {
                return parts.Last();
            }
            else
            {
                return null; // 如果分割后没有部分，返回null
            }


        }

        // 解析RevisionNumber
        private static string GetRevisionNumber(string fileName)
        {
            string pattern = @"_R(\d+)_";
            System.Text.RegularExpressions.Match match = Regex.Match(fileName, pattern);
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            return string.Empty;
        }

        // 解析Discipline
        private static string GetDiscipline(string fileName)
        {
            string pattern = @"_([A-Z]{2})_";
            System.Text.RegularExpressions.Match match = Regex.Match(fileName, pattern);
            if (match.Success)
            {
                string disciplineCode = match.Groups[1].Value;
                switch (disciplineCode)
                {
                    case "WL":
                        return "WL - Structural Precast External Walls";
                    case "BR":
                        return "BR - Structural Precast Staircase Beams";
                    case "CU":
                        return "CU- Structural Column";
                    case "HC":
                        return "HC - Structural Precast Hollow Core Slabs";
                    case "LB":
                        return "LB - Structural Precast Beams";
                    case "PS":
                        return "PS - Structural Precast Staircase";
                    case "SD":
                        return "SD - Structural Precast Scupper Drain";
                    case "SP":
                        return "SP - Structural Precast Solid Walls";
                    case "SS":
                        return "SS - Structural Precast Solid Slabs";
                    case "ST":
                        return "ST - Structural Precast Staircase";
                    case "VL":
                        return "VL - Structural Precast Internal Walls";
                    case "VP":
                        return "VP - Structural Precast Parapet Wall";
                    //注意PS有两种含义：PS - Structural Precast Scupper Drain和PS - Structural Precast Staircase
                    //注意PS有两种含义：SP - Structural Precast Solid Walls和SP - Structural Precast Parapet Walls

                    default:
                        return disciplineCode;
                }
            }
            return string.Empty;
        }

        // 解析Levels
        private static string GetLevels(string fileName)
        {
            string[] parts = fileName.Split('_');
            if (parts.Length > 3) // 确保有足够的段
            {
                string levelCode = parts[3]; // 取第4部分，也就是序号3
                switch (levelCode)
                {
                    case "B01":
                        return "B01 - Basement Level 01";
                    case "B02":
                        return "B01 - Basement Level 02";
                    case "B03":
                        return "B01 - Basement Level 04";
                    case "B04":
                        return "B01 - Basement Level 04";
                    case "BFL":
                        return "BFL - Building Foundations Level";
                    case "BGF":
                        return "BGF - Building Ground Floor";
                    case "BRF":
                        return "BRF - Building Roof";
                    case "ESP":
                        return "ESP - External Spaces for Villas";
                    case "IAG":
                        return "IAG - Above ground infrastructure utilities";
                    case "INF":
                        return "INF - Infrastructure utilities : used only when(IAG, IUG) codes are not applicable";
                    case "IUG":
                        return "IUG - Underground infrastructure utilities";
                    case "L01":
                        return "L01 - Level 01";
                    case "L02":
                        return "L02 - Level 02";
                    case "L03":
                        return "L03 - Level 03";
                    case "L04":
                        return "L04 - Level 04";
                    case "L05":
                        return "L05 - Level 05";
                    case "L06":
                        return "L06 - Level 06";
                    case "L07":
                        return "L07 - Level 07";
                    case "LGF":
                        return "LGF - Lower Ground Floor";
                    case "LRF":
                        return "LRF - Lower Roof";
                    case "P01":
                        return "P01 - Podium 01";
                    case "P02":
                        return "P02 - Podium 02";
                    case "P03":
                        return "P03 - Podium 03";
                    case "UGF":
                        return "Upper Ground Floor";
                    case "URF":
                        return "URF - Upper Roof";
                    default:
                        return levelCode;
                }
            }
            return string.Empty;
        }

        // 清理无效的文件名字符
        private static string CleanInvalidFileNameChars(string filename)
        {
            if (string.IsNullOrEmpty(filename)) return "";
            return string.Join("_", filename.Split(Path.GetInvalidFileNameChars()));
        }
    }
}