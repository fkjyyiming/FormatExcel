using System.Collections.Generic;
using System.Windows;

namespace FormatExcel
{
    /// <summary>
    /// CompareResultWindow.xaml 的交互逻辑
    /// </summary>
    public partial class CompareResultWindow : Window
    {
        public CompareResultWindow()
        {
            InitializeComponent();
        }
        public void SetResult(int totalCompared, int totalDiff, List<string> diffPdf, List<string> diffDwg)
        {
            txtReportSummary.Text = $"共有 {totalCompared} 组文件进行了对比。\n" +
                                      $"共有 {totalDiff/2} 组文件命名有差异。";

            lstPdfFiles.ItemsSource = diffPdf;
            lstDwgFiles.ItemsSource = diffDwg;
        }

        // PDF 列表复制文本
        private void CopyPdfText_Click(object sender, RoutedEventArgs e)
        {
            if (lstPdfFiles.SelectedItem != null)
            {
                Clipboard.SetText(lstPdfFiles.SelectedItem.ToString());
                //MessageBox.Show("已复制");
            }
        }

        // DWG 列表复制文本
        private void CopyDwgText_Click(object sender, RoutedEventArgs e)
        {
            if (lstDwgFiles.SelectedItem != null)
            {
                Clipboard.SetText(lstDwgFiles.SelectedItem.ToString());
                //MessageBox.Show("已复制");
            }
        }
    }

}
