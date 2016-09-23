#include "CExcelOperation.h"

CExcelOperation::CExcelOperation()
{
	m_lpDisp = NULL;
	m_lMaxColNum = -1;
	m_lMaxRowNum = -1;
	m_lStartCol = -1;
	m_lStartRow = -1; 
	m_pfExcelDealCB = NULL;
}
CExcelOperation::~CExcelOperation()
{
}

int CExcelOperation::openExcelFile(CString strFileName)
{
	int iRet = -1; 
	m_strErrMsg = "No Error";
	CRange rangeUsedCells;
	CRange rangeTempCells;
	COleVariant covResult;
	COleVariant covOption((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	do 
	{
		if (strFileName.IsEmpty())
		{
			m_strErrMsg = "Error, strFileName is empty!";
			break;
		}

		if (!m_appExcelServer.CreateDispatch("Excel.Application")) {
			m_strErrMsg = "Error, open Excel Application Server fail!";
			break;
		}
		m_wbsExcelBooks.AttachDispatch(m_appExcelServer.get_Workbooks());
        
		m_lpDisp = m_wbsExcelBooks.Open(strFileName,covOption, covOption, covOption, covOption, covOption, 
			                            covOption, covOption, covOption, covOption, covOption, covOption, 
							            covOption, covOption, covOption);
		if (m_lpDisp == NULL)
		{
			break;
		}
		m_wbExcelBook.AttachDispatch(m_lpDisp);
		m_lpDisp = m_wbExcelBook.get_ActiveSheet();
		m_wsSheet.AttachDispatch(m_lpDisp);
		m_rangeBasicCells.AttachDispatch(m_wsSheet.get_Cells());

		rangeUsedCells.AttachDispatch(m_wsSheet.get_UsedRange());
		rangeTempCells.AttachDispatch(rangeUsedCells.get_Columns());
		m_lMaxColNum = rangeTempCells.get_Count();
		rangeTempCells.AttachDispatch(rangeUsedCells.get_Rows());
		m_lMaxRowNum = rangeTempCells.get_Count();
		m_lStartCol = rangeUsedCells.get_Column();
		m_lStartRow = rangeUsedCells.get_Row();

		rangeUsedCells.ReleaseDispatch();
		rangeTempCells.ReleaseDispatch();

		iRet = 0;
	} while(0);

	return iRet;
}

int CExcelOperation::closeExcelFile()
{
	m_rangeBasicCells.ReleaseDispatch();
	m_wsSheet.ReleaseDispatch();
	m_wbExcelBook.ReleaseDispatch();
	m_wbsExcelBooks.Close();
	m_wbsExcelBooks.ReleaseDispatch();
	m_appExcelServer.Quit();
	m_appExcelServer.ReleaseDispatch();
	return 0;
}

int CExcelOperation::startDealExcel()
{
	int iRet = -1; 
	m_strErrMsg = "No Error";

	do 
	{
		if (m_pfExcelDealCB == NULL)
		{
			m_strErrMsg = "Error: CallBack is NULL!";
			break;
		}
		m_pfExcelDealCB(this);
		//AfxBeginThread((AFX_THREADPROC)m_pfExcelDealCB,this,THREAD_PRIORITY_NORMAL);

		iRet = 0;
	} while(0);
	return iRet;
}