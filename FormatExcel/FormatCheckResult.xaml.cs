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
            ResultTextBlock.Text = $" {issueCount} issues in total, as listed below.";
            IssuesListView.ItemsSource = issues;
        }


        // 右键菜单逻辑
        private void IssuesListView_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is TextBlock textBlock && textBlock.DataContext != null)
            {
                IssuesListView.ContextMenu.Tag = textBlock.DataContext;
                foreach (MenuItem item in IssuesListView.ContextMenu.Items)
                {
                    item.Tag = textBlock;
                }
            }
        }

        // 复制文件名称
        private void CopyFileNameMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.Tag is TextBlock textBlock)
            {
                var dataContext = textBlock.DataContext;
                var fileName = dataContext.GetType().GetProperty("FileName")?.GetValue(dataContext)?.ToString();
                if (!string.IsNullOrEmpty(fileName))
                {
                    Clipboard.SetText(fileName);
                    //MessageBox.Show($"已复制文件名称: {fileName}");
                }
            }
        }

        // 复制问题类型
        private void CopyIssueTypeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.Tag is TextBlock textBlock)
            {
                var dataContext = textBlock.DataContext;
                var issueType = dataContext.GetType().GetProperty("IssueType")?.GetValue(dataContext)?.ToString();
                if (!string.IsNullOrEmpty(issueType))
                {
                    Clipboard.SetText(issueType);
                    //MessageBox.Show($"已复制问题类型: {issueType}");
                }
            }
        }
    }
}
