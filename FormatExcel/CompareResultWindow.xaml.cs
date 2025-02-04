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
            txtReportSummary.Text = $"A total of {totalCompared} pairs of files were compared.\n" +
                                      $"A total of {totalDiff / 2} pairs of files exhibited naming discrepancies.";

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
