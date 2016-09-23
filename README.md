# MFC-Excel
MFC表格操作模板

## 使用说明

0.1 在项目中添加以下两个文件

CExcelOperation.h
CExcelOperation.cpp

1.1 右键CExcelOperation.cpp，选择属性

1.2 C/C++ -> 预编译头 ->不使用预编译头

2.1 选中MFC项目中，右键添加类，选择TypeLib  

2.2 来源 -> 注册表，可用库类型 -> Microsoft Excel xx

2.3 接口选择以下七项

_Application
_Workbook
_Worksheet
Workbooks
Worksheets
Sheets
Rangs

2.4 将新生成的七个头文件中的"#import "...EXCEL.EXE" no_namespace"去掉

2.5 将CRange.h中的VARIANT DialogBox()改成VARIANT _DialogBox()
