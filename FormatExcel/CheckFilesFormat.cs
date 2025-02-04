using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FormatExcel
{
    public class CheckFilesFormat
    {
        public class FileIssue
        {
            public string FileName { get; set; }
            public string IssueType { get; set; }
        }

        public static (int IssueCount, List<FileIssue> Issues) CheckFiles(string folderPath, string formatEx)
        {
            // formatEx 为文件扩展名，如".pdf"，".dwg"等
            // 文件命名实例：00103_CHE_VL1_BGF_SHD_CU_002001_R02_V1-SC001.pdf 或者 00103_CHE_VL1_BGF_SHD_CU_002001_R02_V1-SC001.dwg等类似结构

            //定义问题数量和问题列表
            int issueCount = 0;
            List<FileIssue> issues = new List<FileIssue>();

            //检查文件夹是否存在
            if (!Directory.Exists(folderPath))
            {
                return (issueCount, issues); // Or throw an exception if the folder doesn't exist
            }

            //获取文件夹下所有指定格式的文件
            string[] files = Directory.GetFiles(folderPath, "*" + formatEx, SearchOption.TopDirectoryOnly);
            foreach (string filePath in files)
            {

                //判断条件1

                //  获取真实文件名（不带拓展名）
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
                string trueName = fileNameWithoutExtension.Split('.').FirstOrDefault();
               
                
                //使用下划线 "_" 分割文件名
                string[] parts = trueName.Split('_');



                // 错误情况1: 文件被_分割后的部分数量不等于8
                if (parts.Length != 9)
                {
                    issueCount++;
                    //利用辅助方法 UpdateIssueList 更新 FileIssue 列表 （该方法为static）
                    UpdateIssueList(issues, fileNameWithoutExtension, "File length or delimiter error. Please modify and recheck.");

                }

                else
                {
                    //注意原先在上一层作用域，但是若文件分隔不对，会报错卡死，因此将文件分隔检查作为第一步

                    var assetCode = parts[0];
                    var orignator = parts[1];
                    var unitType = parts[2];
                    var level = parts[3];
                    var drawingType = parts[4];
                    var elementType = parts[5];
                    var elementNum = parts[6];
                    var version = parts[7];
                    var elementDetail = parts[8];
                    // 错误情况2：纯数字部分或版本号长度错误
                    if ((assetCode.Length != 5) || (elementNum.Length != 6) || (version.Length != 3))
                    {
                        issueCount++;
                        UpdateIssueList(issues, fileNameWithoutExtension, "Number or version section length error.");
                    }
                    // 错误情况3：构件编号不符合规范
                    // 00103_CHE_VL1_L02_SHD_HC_002150_R03_V1-2HC1500   2HC1500  15 150厚的板 序号0  对应002150  15 是厚度 0 是序号
                    // 00103_CHE_VL1_L02_SHD_HC_002260_R02_V1-2HC2510   2HC2510  25  250厚的板 序号10  对应002260 （应该为00225x)但是超号，+1成002260


                    var (ele_Level, ele_Category, ele_Num) = GetStringParts(elementDetail);
                    var a = ele_Level;
                    var b = ele_Category;
                    var c = ele_Num;

                    // 错误情况3.1 ：对于长（HC、SS）、短（ST）之外的类型（最后一段的长度为5，其中为2位字母，3位数字），直接将前面纯数字后三位进行对比，若不一致则报错
                    if (ele_Category != "HC" && ele_Category != "SS" && ele_Category != "ST")
                    {
                        //取纯数字部分的后三位
                        var numFrontPart = elementNum.Substring(elementNum.Length - 3, 3);

                        if ((elementDetail.Split('-').ToList().Last().Length) != 5)
                        {
                            issueCount++;
                            UpdateIssueList(issues, fileNameWithoutExtension, "Final numbering length error.");

                        }
                        else if (ele_Num != numFrontPart)
                        {
                            issueCount++;
                            UpdateIssueList(issues, fileNameWithoutExtension, "Inconsistent component numbering.");
                        }
                        else if (ele_Level != elementNum[3].ToString())
                        {
                            issueCount++;
                            UpdateIssueList(issues, fileNameWithoutExtension, "Elevation mismatch.");
                        }
                        else
                        {
                            continue;
                        }

                    }

                    // 错误情况3.2 ：对于短（ST）类型（总共长度为4，其中2位字母，2位数字，最后一段无楼层信息），直接将前面纯数字后两位进行对比，若不一致则报错
                    if (ele_Category == "ST")
                    {
                        var numFrontPart = elementNum.Substring(elementNum.Length - 2, 2);
                        if ((elementDetail.Split('-').ToList().Last().Length) != 4)
                        {
                            issueCount++;
                            UpdateIssueList(issues, fileNameWithoutExtension, "Final numbering length error.");

                        }
                        else if (ele_Num != numFrontPart)
                        {
                            issueCount++;
                            UpdateIssueList(issues, fileNameWithoutExtension, "Inconsistent component numbering.");
                        }
                        else
                        {
                            continue;
                        }

                    }

                    // 错误情况3.3 ：对于长类型，由于进位情况比较特殊，所以分两种情况进行判断
                    if (ele_Category == "HC" || ele_Category == "SS")
                    {

                        var numFrontPart = elementNum.Substring(elementNum.Length - 2, 1);

                        //最后一段的倒数第一位数字
                        var keyMark_R1 = elementDetail.Substring(elementDetail.Length - 1, 1);
                        //最后一段的倒数第二位数字
                        var keyMark_R2 = elementDetail.Substring(elementDetail.Length - 2, 1);
                        //最后一段的倒数第三位数字
                        var keyMark_R3 = elementDetail.Substring(elementDetail.Length - 3, 1);
                        //最后一段的倒数第四位数字
                        var keyMark_R4 = elementDetail.Substring(elementDetail.Length - 4, 1);

                        //前面序列号的后三位数字
                        var seqMark = elementNum.Substring(elementNum.Length - 3);





                        //长度固定为7位，若不是则报错
                        if ((elementDetail.Split('-').ToList().Last().Length) == 7)
                        {
                            if (keyMark_R2 == "0")
                            {

                                var seqFour = seqMark.Insert(seqMark.Length - 1, "0");
                                if (ele_Num != seqFour)
                                {
                                    issueCount++;
                                    UpdateIssueList(issues, fileNameWithoutExtension, "Inconsistent component numbering.");
                                }
                            }
                            else
                            {
                                //最后一段倒数第二位和倒数第三位相加，再转换位字符串，与前面序列号的倒数第三位数字进行对比

                                var keyMark = (int.Parse(keyMark_R2) + int.Parse(keyMark_R3)).ToString();
                                var lastPart = keyMark_R4 + keyMark + keyMark_R1;

                                if (seqMark != lastPart)
                                {
                                    issueCount++;
                                    UpdateIssueList(issues, fileNameWithoutExtension, "Inconsistent component numbering.");
                                }

                            }

                        }
                        else
                        {
                            issueCount++;
                            UpdateIssueList(issues, fileNameWithoutExtension, "Final numbering length error.");
                        }

                    }
                


                }


            }






            ////对于长（HC、SS）、短（ST）之外的类型（最后一段的长度为5），直接将四位数字进行对比
            //if (elementDetail.Length == 5)
            //{
            //    var numFront = elementNum;
            //    var numFrontPart = numFront.Substring(numFront.Length - 1, 3);
            //    //对于长（HC、SS）、短（ST）之外的类型，直接将四位数字进行对比
            //    if (ele_Category != "HC" || ele_Category != "SS" || ele_Category != "ST")
            //    {
            //        if (ele_Num != numFrontPart)
            //        {
            //            issueCount++;
            //            UpdateIssueList(issues, fileNameWithoutExtension, "构件编号前后不一致");
            //        }
            //        else
            //        {
            //            continue;
            //        }

            //    }
            //}



            return (issueCount, issues);
        }

        // 辅助方法：更新 FileIssue 列表
        private static void UpdateIssueList(List<FileIssue> issues, string fileName, string issueType)
        {
            FileIssue existingIssue = issues.FirstOrDefault(x => x.FileName == fileName);

            //如果不存在，则添加新的 FileIssue
            if (existingIssue == null)
            {
                issues.Add(new FileIssue { FileName = fileName, IssueType = issueType });
            }
            //如果存在，则更新 IssueType
            else
            {
                existingIssue.IssueType += "," + issueType;
            }
        }


        /// <summary>
        /// 获取文件名最后一部分的字符串列表，如"V1-SC001"或"V1-2HC1500"
        /// 以-拆分后，分为两种情况
        /// 若第二部分长度为5，则分隔为长度为2的字母和长度为3的数字，并将3位数字的第一个数字作为楼层
        /// 若第二部分长度为7，则分隔为长度为1的数字（楼层）、长度为3的字母（构件类型）、长度为4的字母（构件编号）
        /// </summary>
        /// <param name="lastPart">输入为文件名的最后一部分，如"V1-SC001"或"V1-2HC1500"</param>
        /// <returns></returns>
        private static (string, string, string) GetStringParts(string lastPart)
        {
            var lastParts = lastPart.Split('-').ToList();
            var lastNum = lastParts.Last();
            return SperateNumExtractThreeParts(lastNum);
        }

        /// <summary>
        /// 获取文件名最后一部分的字符串列表，如"SC001"或"2HC1500"，返回楼层、构件类型、构件编号
        /// </summary>
        /// <param name="lastParts"></param>
        /// <returns></returns>
        private static (string, string, string) SperateNumExtractThreeParts(string lastParts)
        {
            string pattern = @"^(\d*)([A-Z]+)(\d+)$";
            var match = Regex.Match(lastParts, pattern, RegexOptions.IgnoreCase);

            // 提取第一个数字
            string firstNumber = match.Groups[1].Value;

            // 提取字母部分
            string letters = match.Groups[2].Value;

            // 提取最后数字部分
            string lastNumber = match.Groups[3].Value;

            // 
            if (firstNumber.Length > 0)
            {
                return (firstNumber, letters, lastNumber);
            }
            else if (lastNumber.Length == 2)
            {
                return ("", letters, lastNumber);
            }
            else
            {
                return (lastNumber[0].ToString(), letters, lastNumber);
            }

        }

    }
}
