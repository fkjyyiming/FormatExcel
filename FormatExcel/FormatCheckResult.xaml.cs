using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace FormatExcel
{
    /// <summary>
    /// FormatCheckResult.xaml 的交互逻辑
    /// </summary>
    public partial class FormatCheckResult : Window
    {
        public FormatCheckResult()
        {
            InitializeComponent();
        }

        internal void SetResult(int issueCount, List<CheckFilesFormat.FileIssue> issues)
        {
            ResultTextBlock.Text = $"共有 {issueCount} 个文件有问题，如下列表";
            IssuesListView.ItemsSource = issues;
        }
    }
}
