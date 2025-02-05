# 1 项目目的
领域：建筑工程、BIM、装配式建筑  
语言/框架/程序包：C#/WPF/NPOI（EXCEL的生成）  
本项目适用沙特Sedra项目文件规范化管理，每次完成装配式户型的加工图设计的时候，需要提交MIDP表格，该表格的整理相当费时费力。  
因此通过C#结合EXCEL样板文件，依据CAD 或者 PDF文件夹中的文件，自动提取关键信息，进行相关表格的创建  
将相关代码传至仓库，做记录。同时若有相关需求的大神，可以随意自取。   
  
软件界面  
![image](https://github.com/user-attachments/assets/e62bb3ee-6d95-4e61-81f5-f3249e9e7a05)  


# 2 主要功能
## 2.1 路径选取
WPF界面下，可以通过按钮进行选取，也可以将路径直接粘贴至文本框

## 2.2 文件对比
CAD 和 PDF文件夹内的主要文件，应当是除了后缀名后完全一致的，将对比的结果以列表的形式进行呈现。  相关结果可右键复制内容。

## 2.3 文件查错
项目对于.pdf和.dwg文件的命名方式有严格的要求，例如资产代码、序列号、户型信息等，做了几种规则检查，以适配大部分的情况。  
将检查结果以列表的形式进行呈现。  相关结果可右键复制内容。

## 2.4 新/旧MIDP表格生成
项目新/旧MIDP表格同时都在用，因此适配两种情况。  
界面的文本内容是缺省项，可以不填。  
其他MIDP表格中的内容都依据文件名的各字段自动生成。  
表格生成引用了NPOI程序包，利用其生成EXCEL  
同时添加了MVS Installer Projects 扩展，可以对程序进行打包。  
使用方可以像安装普通软件一样进行安装操作。  

# 3 增加英文版
为方便外籍使用，在分支中增加纯英文版本  
![0500577baf36379e2e1aef17ea2def9](https://github.com/user-attachments/assets/00300a70-7788-4f97-a553-84dfc7694978)
