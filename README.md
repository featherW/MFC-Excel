# MFC-Excel
MFC表格操作模板

## 使用说明

0.1 在项目中添加以下两个文件

CExcelOperation.h

CExcelOperation.cpp

---


1.1 右键CExcelOperation.cpp，选择属性

1.2 C/C++ -> 预编译头 ->不使用预编译头

---


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


---

注意事项

1 如果在写操作后出现程序退出后，Excel进程没有退出。

可以在int CExcelOperation::closeExcelFile()试试添加m_wbExcelBook.Save();

```
int CExcelOperation::closeExcelFile()
{
	m_rangeBasicCells.ReleaseDispatch();
	m_wsSheet.ReleaseDispatch();
	m_wbExcelBook.Save();
	m_wbExcelBook.ReleaseDispatch();
	m_wbsExcelBooks.Close();
	m_wbsExcelBooks.ReleaseDispatch();
	m_appExcelServer.Quit();
	m_appExcelServer.ReleaseDispatch();
	return 0;
}
```
