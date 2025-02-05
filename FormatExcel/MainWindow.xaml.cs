﻿using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Windows;

namespace FormatExcel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private string pdfFolderPath;
        private string dwgFolderPath;
        private string templatePath;
        private string templatePath_NewMIDP;

        public MainWindow()
        {
            InitializeComponent();

            // 设置左侧默认值
            TxtSheetSize.Text = "A1";
            TxtScale.Text = "1:50";
            TxtDrawingType.Text = "General";
            TxtDiscipline2.Text = "PRECAST";
            // 设置右侧默认值
            TxtDesignStage.Text = "08 - PDD - Product Detailed Design";
            TxtCategory.Text = "SHD - Shop Drawings";
            TxtToCompany.Text = "East Consulting Engineering Company";
            TxtZones.Text = "";
            // 设置模板路径--旧的MIDP模版
            templatePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template", "ExcelTemplate(Don not remove and modify).xlsx");

            // 设置模版路径--新的MIDP模版
            templatePath_NewMIDP = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template", "ExcelTemplate_NewMIDP(Don not remove and modify).xlsx");


        }

        // 按钮1: 对比PDF和DWG
        private void BtnComparePDFDWG_Click(object sender, RoutedEventArgs e)
        {
            // 1. 检查路径是否为空
            if (string.IsNullOrEmpty(TxtPDFPath.Text) || string.IsNullOrEmpty(TxtDWGPath.Text))
            {
                System.Windows.MessageBox.Show("请先选择对比文件夹\nPlease select the comparison folders first.", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            pdfFolderPath = TxtPDFPath.Text;
            dwgFolderPath = TxtDWGPath.Text;

            // 调用 FileComparer 并获取结果
            var (totalCompared, totalDiff, diffPdf, diffDwg) = FileComparer.CompareFiles(pdfFolderPath, dwgFolderPath);

            // 创建并显示结果窗口
            CompareResultWindow resultWindow = new CompareResultWindow();
            resultWindow.SetResult(totalCompared, totalDiff, diffPdf, diffDwg);
            resultWindow.ShowDialog();
        }


        // 按钮2: 生成标准化表格
        private void BtnGenerateOLDTable_Click(object sender, RoutedEventArgs e)
        {
            // 1. 检查PDF路径是否选择
            if (string.IsNullOrEmpty(TxtPDFPath.Text))
            {
                System.Windows.MessageBox.Show("请至少保证选择PDF文件夹\nPlease ensure at least the PDF folder is selected.", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!File.Exists(templatePath))
            {
                System.Windows.MessageBox.Show("模板文件未找到!\nTemplate file not found.", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            pdfFolderPath = TxtPDFPath.Text;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "另存为标准化表格(Save as standardized table)";
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                string savePath = saveFileDialog.FileName;

                if (!string.IsNullOrEmpty(savePath))
                {
                    try
                    {
                        ExcelGenerator.GenerateExcelReport(pdfFolderPath, TxtDesignStage.Text, TxtCategory.Text, TxtToCompany.Text, TxtZones.Text, templatePath, savePath);
                        System.Windows.MessageBox.Show("标准化表格生成完成！\n ", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show(ex.Message, "错误(Error!)", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("请选择一个有效的保存路径\nPlease select a valid save path.", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // 按钮3: 选择PDF路径 (选择文件夹)
        private void BtnSelectPDFPath_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true; // 设置为文件夹选择模式
            dialog.Title = "选择PDF文件夹(Please select the PDF folder)";

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                pdfFolderPath = dialog.FileName; // 获取选择的文件夹路径
                TxtPDFPath.Text = pdfFolderPath;  // 显示在 TextBox 中
            }
        }

        // 按钮4: 选择DWG路径 (选择文件夹)
        private void BtnSelectDWGPath_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true; // 设置为文件夹选择模式
            dialog.Title = "选择DWG文件夹(Please select the DWG folder)";

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                dwgFolderPath = dialog.FileName; // 获取选择的文件夹路径
                TxtDWGPath.Text = dwgFolderPath; // 显示在 TextBox 中
            }
        }

        private void CheckDWGFormat_Click(object sender, RoutedEventArgs e)
        {
            // 1. 检查路径是否为空
            if (string.IsNullOrEmpty(TxtDWGPath.Text))
            {
                System.Windows.MessageBox.Show("请先选择DWG文件夹\nPlease select the DWG folder first.", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            dwgFolderPath = TxtDWGPath.Text;

            // 调用 FileComparer 并获取结果
            var (issueCount, issues) = CheckFilesFormat.CheckFiles(dwgFolderPath,".dwg");

            // 创建并显示结果窗口
            FormatCheckResult resultWindow = new FormatCheckResult();
            resultWindow.SetResult(issueCount, issues);
            resultWindow.ShowDialog();
        }

        private void CheckPDFFormat_Click(object sender, RoutedEventArgs e)
        {
            // 1. 检查路径是否为空
            if (string.IsNullOrEmpty(TxtPDFPath.Text))
            {
                System.Windows.MessageBox.Show("请先选择PDF文件夹\nPlease select the PDF folder first.", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            pdfFolderPath = TxtPDFPath.Text;

            // 调用 FileComparer 并获取结果
            var (issueCount, issues) = CheckFilesFormat.CheckFiles(pdfFolderPath, ".pdf");

            // 创建并显示结果窗口
            FormatCheckResult resultWindow = new FormatCheckResult();
            resultWindow.SetResult(issueCount, issues);
            resultWindow.ShowDialog();
        }

        private void BtnGenerateNEWTable_Click(object sender, RoutedEventArgs e)
        {

            // 1. 检查PDF路径是否选择
            if (string.IsNullOrEmpty(TxtPDFPath.Text))
            {
                System.Windows.MessageBox.Show("请至少保证选择PDF文件夹\nPlease ensure at least the PDF folder is selected.", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!File.Exists(templatePath_NewMIDP))
            {
                System.Windows.MessageBox.Show("模板文件未找到!\nTemplate file not found.", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            pdfFolderPath = TxtPDFPath.Text;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "另存为标准化表格(Save as standardized table--new MIDP)";
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                string savePath = saveFileDialog.FileName;

                if (!string.IsNullOrEmpty(savePath))
                {
                    try
                    {
                        ExcelGenerator.GenerateExcelReportNewMIDP(pdfFolderPath, TxtSheetSize.Text, TxtScale.Text, TxtDrawingType.Text, TxtDiscipline2.Text, templatePath_NewMIDP, savePath);
                        System.Windows.MessageBox.Show("标准化表格生成完成！\n ", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show(ex.Message, "错误(Error!)", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("请选择一个有效的保存路径\nPlease select a valid save path.", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }



        }
    }
}