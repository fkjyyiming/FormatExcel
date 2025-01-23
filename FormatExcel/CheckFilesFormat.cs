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

        public static (int IssueCount, List<FileIssue> Issues) CheckFiles(string folderPath)
        {

            int issueCount = 0;
            List<FileIssue> issues = new List<FileIssue>();

            if (!Directory.Exists(folderPath))
            {
                return (issueCount, issues); // Or throw an exception if the folder doesn't exist
            }

            string[] files = Directory.GetFiles(folderPath);

            foreach (string file in files)
            {
                // TODO: Implement your file checking logic here
                // Example: Check if the file name contains a specific pattern
                // if (!Path.GetFileName(file).Contains("_compliant_"))
                // {
                //     issueCount++;
                //     issues.Add(new FileIssue { FileName = Path.GetFileName(file), IssueType = "Missing _compliant_ in name" });
                // }

                // Placeholder for demonstration: Let's say every other file has an issue
                if (Array.IndexOf(files, file) % 2 != 0)
                {
                    issueCount++;
                    issues.Add(new FileIssue { FileName = Path.GetFileName(file), IssueType = "Placeholder Issue" });
                }
            }

            return (issueCount, issues);
        }
    }
}
