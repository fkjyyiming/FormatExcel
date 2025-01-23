using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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

                // 00103_CHE_VL1_L02_SHD_HC_002150_R03_V1-2HC1500   2HC1500  15 150厚的板 序号0  对应002150  15 是厚度 0 是序号
                // 00103_CHE_VL1_L02_SHD_HC_002260_R02_V1-2HC2510   2HC2510  25  250厚的板 序号10  对应002260 （应该为00225x)但是超号，+1成002260
                //判断条件1

                //  获取真实文件名（不带拓展名）
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);

                //使用下划线 "_" 分割文件名
                string[] parts = fileNameWithoutExtension.Split('_');
                var assetCode = parts[0];
                var orignator = parts[1];
                var unitType = parts[2];
                var level = parts[3];
                var drawingType = parts[4];
                var elementType = parts[5];
                var elementNum = parts[6];
                var version = parts[7];
                var elementDetail = parts[8];

                // 定义错误情况 (至少三种)
                // 错误情况 1: 文件被_分割后的部分数量不等于8
                if (parts.Length != 9)
                {
                    issueCount++;
                    UpdateIssueList(issues, fileNameWithoutExtension, "文件长度或分隔_错误");
                }
                if ((parts[0].Length != 5) || (parts[6].Length != 6) || (parts[7].Length != 3))
                {
                    issueCount++;
                    UpdateIssueList(issues, fileNameWithoutExtension, "纯数字部分或版本号长度错误");
                }





            }

            return (issueCount, issues);
        }

        // 辅助方法：更新 FileIssue 列表
        private static void UpdateIssueList(List<FileIssue> issues, string fileName, string issueType)
        {
            FileIssue existingIssue = issues.FirstOrDefault(x => x.FileName == fileName);

            if (existingIssue == null)
            {
                issues.Add(new FileIssue { FileName = fileName, IssueType = issueType });
            }
            else
            {
                existingIssue.IssueType += "," + issueType;
            }
        }
    }
}
