#ifndef __CEXCEL_OPERATION_H__
#define __CEXCEL_OPERATION_H__

/* CExcelOperation : MFC操作Excel表格
 * by Bai Yingfei 2016.02.26
 *
 * 使用说明
 * 1.1 右键CExcelOperation.cpp，选择属性
 * 1.2 C/C++ -> 预编译头 ->不使用预编译头
 *
 * 2.1 选中MFC项目中，右键添加类，选择TypeLib  
 * 2.2 来源 -> 注册表，可用库类型 -> Microsoft Excel xx
 * 2.3 接口选择以下七项
 * _Application
 * _Workbook
 * _Worksheet
 * Workbooks
 * Worksheets
 * Sheets
 * Rangs
 * 2.4 将新生成的七个头文件中的"#import "...EXCEL.EXE" no_namespace"去掉
 * 2.5 将CRange.h中的VARIANT DialogBox()改成VARIANT _DialogBox()
 */

#include <iostream>
#include "stdafx.h"
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CSheets.h"
#include "CRange.h"
using std::string;

typedef void (*ExcelDealCB)(void*);

class CExcelOperation
{
public:
	CExcelOperation();
	~CExcelOperation();

    /*************** 
     * Funciton:  openExcelFile 
     * Describe:  打开一个Excel表格文件.
     * Parameter: strFileName(in) 表格文件的路径.
     * Return:    返回0成功，否则失败，使用getErrMsg()可以查看错误信息.
     ***************/ 
	int openExcelFile(CString strFileName);

    /*************** 
     * Funciton:  closeExcelFile
     * Describe:  关闭Excel表格文件
     * Parameter: 
     * Return:    0
     ***************/ 
	int closeExcelFile();

	/*************** 
     * Funciton:  setDealCB  
     * Describe:  设置操作回调到m_pfExcelDealCB中.
     * Parameter: excelDealCB(in) 操作表格的回调函数. 
     * Return:    
     ***************/ 
	void setDealCB(ExcelDealCB excelDealCB) {m_pfExcelDealCB = excelDealCB;};

    /*************** 
     * Funciton:  startDealExcel
     * Describe:  对Excel表格进行操作，需要先调用setDealCB()设置操作回调.
     * Parameter: 
     * Return:    返回0成功，否则失败，使用getErrMsg()可以查看错误信息.
     ***************/ 
	int startDealExcel();

    /*************** 
     * Funciton:  clearErrMsg
     * Describe:  清空错误信息m_strErrMsg.
     * Parameter: 
     * Return:    
     ***************/ 
	void clearErrMsg() {m_strErrMsg.clear();};

    /*************** 
     * Funciton:  getErrMsg  
     * Describe:  得到错误信息的内容.
     * Parameter: 
     * Return:    错误信息m_strErrMsg.
     ***************/ 
	string getErrMsg() {return m_strErrMsg;};

    /*************** 
     * Funciton:  getCell  
     * Describe:  得到可以操作表格的类.
     * Parameter: 
     * Return:    可操作表格的类m_rangeBasicCells.
     ***************/ 
	CRange getCell() {return m_rangeBasicCells;};

    /*************** 
     * Funciton:  getMaxColNum  
     * Describe:  得到表格的最大列数.
     * Parameter: 
     * Return:    表格的最大列数.
     ***************/ 
	long getMaxColNum() {return m_lMaxColNum;};

    /*************** 
     * Funciton:  getMaxColNum  
     * Describe:  得到表格的最大行数.
     * Parameter: 
     * Return:    表格的最大行数.
     ***************/ 
	long getMaxRowNum() {return m_lMaxRowNum;};

    /*************** 
     * Funciton:  getStartCol
     * Describe:  得到表格的起始列号.  
     * Parameter: 
     * Return:    表格的起始列号.
     ***************/ 
	long getStartCol() {return m_lStartCol;};

    /*************** 
     * Funciton:  getStartRow  
     * Describe:  得到表格起始行号.
     * Parameter: 
     * Return:    表格的起始行号.
     ***************/ 
	long getStartRow() {return m_lStartRow;};

private:
	CApplication m_appExcelServer; //Excel服务器
	CWorkbooks m_wbsExcelBooks;
	CWorkbook m_wbExcelBook; //打开的文档
	LPDISPATCH m_lpDisp;
	CWorksheet m_wsSheet;
	CRange m_rangeBasicCells;
	long m_lMaxColNum;  //表格最大列数
	long m_lMaxRowNum;  //表格最大行数
	long m_lStartCol; //表格起始列
	long m_lStartRow;  //表格起始行
	string m_strErrMsg; //错误信息
	ExcelDealCB m_pfExcelDealCB; //对表格进行操作的回调函数
	
};
#endif
