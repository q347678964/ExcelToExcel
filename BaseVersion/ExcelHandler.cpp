#include "stdafx.h"
#include "ExcelHandler.h"
#include "ExcelLib/OperateExcelFile.h"
#include "resource.h"		// 主符号

#define DEBUG_LOG 0


IllusionExcelFile InExcelOperate;
IllusionExcelFile OutExcelOperate;


CString Mail_SpecialHandler(CString InputMailCString)
{
	CString Ret;

	InputMailCString.Delete(0, InputMailCString.Find(_T(":")) + 1);

	InputMailCString.Delete(0, InputMailCString.Find(_T(" ")) + 1);

	Ret = InputMailCString;

	return Ret;
}

CString Money_SpecialHandler(CString InputMoneyCString)
{
	CString Ret;

	unsigned int SpeicalCharPoint = InputMoneyCString.Find(_T("金"));
	unsigned int CstringLength = InputMoneyCString.GetLength();

	InputMoneyCString.Delete(SpeicalCharPoint, CstringLength);

	Ret = InputMoneyCString;

	return Ret;
}
/*
IN_OUT_CONFIG_S gDefineInOutConfig[MAX_CELL] = {
	{3, 1, CString("0"), 1, Mail_SpecialHandler},\
	{6, 1, CString("0"), 2,Money_SpecialHandler},\
	{9, 2, CString("0"), 3, NULL},\
	{11, 2, CString("0"), 4, NULL},\
	{19, 3, CString("0"), 5, NULL},\
	{23, 2, CString("0"), 6, NULL},\
	{25, 2, CString("0"), 7, NULL},\
	{45, 3, CString("0"), 8, NULL},\
	{49, 2, CString("0"), 9, NULL},\
	{53, 2, CString("0"), 10, NULL},\
	{61, 2, CString("0"), 11, NULL},\
	{9, 3, CString("0"), 16, NULL},\
	{11, 3, CString("0"), 17, NULL},\
};
*/
IN_OUT_CONFIG_S gDefineInOutConfig[MAX_CELL] = {
	{3, 1, CString("0"), 1, Mail_SpecialHandler},\
	{6, 1, CString("0"), 2,Money_SpecialHandler},\
	{9, 2, CString("0"), 3, NULL},\
	{11, 2, CString("0"), 4, NULL},\

	{19, 3, CString("0"), 19, NULL},\
	{21, 3, CString("0"), 20, NULL},\

	{23, 2, CString("0"), 6, NULL},\

	{25, 2, CString("0"), 21, NULL},\
	{25, 3, CString("0"), 22, NULL},\
	{37, 2, CString("0"), 23, NULL},\
	{37, 3, CString("0"), 24, NULL},\

	{45, 2, CString("0"), 25, NULL},\
	{45, 3, CString("0"), 26, NULL},\

	{49, 2, CString("0"), 9, NULL},\
	{53, 2, CString("0"), 10, NULL},\

	{61, 2, CString("0"), 27, NULL},\
	{61, 3, CString("0"), 28, NULL},\

	{9, 3, CString("0"), 16, NULL},\
	{11, 3, CString("0"), 17, NULL},
};

CExcelHandler::CExcelHandler(void)
{
	DebugInfoString = (CString)("");
	CExcelHandler::InputExcelNum = 0;
}

void CExcelHandler::RemoveOutputFile(void)
{
	CFileStatus Fstatus;

	if(CFile::GetStatus(this->OutputExcelPath,Fstatus,NULL)){
		CFile::Remove(this->OutputExcelPath);
	}

	if(CFile::GetStatus(this->OutputExcel_DataCheckPath,Fstatus,NULL)){
		CFile::Remove(this->OutputExcel_DataCheckPath);
	}
	
}

void CExcelHandler::FindInputFile(void)
{

	CString SearchModulePath = CExcelHandler::GetModulePath();
	CString SearchExcelPath;
	CString FileName;
	CFileFind Finder;
	CFileStatus Fstatus;

	this->Printf("[Excel List:]\r\n");

	SearchExcelPath = SearchModulePath + CString("\\..\\Input\\*.xlsx");

	BOOL bWorking = Finder.FindFile(SearchExcelPath);
	while(bWorking)
	{
		bWorking = Finder.FindNextFileW();
		FileName = Finder.GetFileName();

		InputExcelPath[this->InputExcelNum] = SearchModulePath + CString("\\..\\Input\\") + FileName;
		
		/*过滤输出Excel*/
		if(InputExcelPath[this->InputExcelNum] == this->OutputExcel_ModlePath) {
			continue;
		}

		this->PrintfCString(InputExcelPath[CExcelHandler::InputExcelNum] + CString("\r\n"));

		this->InputExcelNum++;

		/*达到最大输入数量，退出查找*/
		if(this->InputExcelNum == MAX_INTPUT_EXCEL) {
			break;
		}
	}

	//DebugInfoString = (CString)("");
	//CExcelHandler::DebugUpdate();

}

void CExcelHandler::DebugUpdate(void)
{
	DebugInfoString+="\r\n";
	
	this->PrintfCString(DebugInfoString);
}

void CExcelHandler::HandlerInputFile(void)
{
	int SheetNumber = 0;
	int InputSheetColumn = 0, InputSheetRow = 0;
	int OutputSheetRow = 0;
	unsigned int i = 0, j = 0;
	int CurSheet = 1;  //Sheet1是数据
	double DoubleData = 0.0;


	/*设置特殊的输入字符串处理函数*/
	for(i = 0;i < this->InputExcelNum; i++) {
		for(j = 0; j < MAX_CELL; j++) {
			gInOutConfig[i][j].Hld = gDefineInOutConfig[j].Hld;
			gInOutConfig[i][j].CurPointCString = CString("0");
			gInOutConfig[i][j].InputColumn = gDefineInOutConfig[j].InputColumn;
			gInOutConfig[i][j].InputRow = gDefineInOutConfig[j].InputRow;
			gInOutConfig[i][j].OutputColumn = gDefineInOutConfig[j].OutputColumn;
		}
	}

	InExcelOperate.InitExcel();

	for(i = 0;i < this->InputExcelNum; i++) {
		InExcelOperate.OpenExcelFile(this->InputExcelPath[i]);

		SheetNumber = InExcelOperate.GetSheetCount();

		
#if 0
		this->Printf("表格数/%u",SheetNumber);
		this->Printf("表格名/%s",InExcelOperate.GetSheetName(CurSheet));

		{			//For Debug
			InExcelOperate.LoadSheet(CurSheet,1);

			InputSheetRow = InExcelOperate.GetRowCount();	//行数量
			this->Printf("行数量/%u",InputSheetRow);

			InputSheetColumn = InExcelOperate.GetColumnCount();	//列数量
			this->Printf("行数量/%u",InputSheetColumn);
			
		}
#else
		InExcelOperate.LoadSheet(CurSheet,1);
		this->Printf("正在处理表格(%u/%u)",i + 1,this->InputExcelNum);
#endif
		this->PrintfCString(this->InputExcelPath[i] + CString("\r\n"));

		//加载Sheet1表格,开始处理数据
		//InExcelOperate.LoadSheet(CurSheet,1);		



		/*读书输入字符串到gInOutConfig[i][j].CurPointCString,  i代表文件， j代表元素*/
		for(j = 0; j < MAX_CELL; j++) {
			CString CellString;
			 CellString = InExcelOperate.GetCellString(gInOutConfig[i][j].InputRow - 1,gInOutConfig[i][j].InputColumn);
			 gInOutConfig[i][j].CurPointCString = CellString;

			if(gInOutConfig[i][j].Hld != NULL) {
				CString ParamCString = gInOutConfig[i][j].CurPointCString;
				gInOutConfig[i][j].CurPointCString = gInOutConfig[i][j].Hld(ParamCString);
			}
			//OutExcelOperate.SetCellString(OutputSheetRow + i, gInOutConfig[j].OutputColumn, gInOutConfig[j].CurPointCString);
		}


		InExcelOperate.CloseExcelFile();
	}

	InExcelOperate.ReleaseExcel();


	
}

void CExcelHandler::CreateOutputFile(void)
{
	int CurSheet = 1;
	int OutputSheetRow = 0;
	int i = 0, j = 0;
	/*下面开始写入Output Excel*/
	OutExcelOperate.InitExcel();
	OutExcelOperate.OpenExcelFile(this->OutputExcel_ModlePath);
	OutExcelOperate.LoadSheet(CurSheet,1);
	OutputSheetRow = OutExcelOperate.GetRowCount();

	OutputSheetRow = 1;

	for(i = 0;i < this->InputExcelNum; i++) {
		for(j = 0; j < MAX_CELL; j++) {
			OutExcelOperate.SetCellString(OutputSheetRow + i + 1, this->gInOutConfig[i][j].OutputColumn, this->gInOutConfig[i][j].CurPointCString);
		}
	}

	OutExcelOperate.SaveasXSLFile(this->OutputExcelPath);
	OutExcelOperate.CloseExcelFile();
	OutExcelOperate.ReleaseExcel();
}

typedef struct _DATA_CHECK_S_ {
	CString ExcelCalcCStringTotalMoney;
	CString ExcelCalcCStringRemainMoney;
	CString InsertCStringTotalMoney;
	CString InsertCStringRemainMoney;
}_DATA_CHECK_S;

typedef struct _DATA_CHECK_DOUBLE_S_ {
	double ExcelCalcCStringTotalMoney;
	double ExcelCalcCStringRemainMoney;
	double InsertCStringTotalMoney;
	double InsertCStringRemainMoney;
}_DATA_CHECK_DOUBLE_S;

_DATA_CHECK_S gDataCheck;
_DATA_CHECK_DOUBLE_S gDoubleDataCheck;

void CExcelHandler::OutputDataCheck(void)
{
	int CurSheet = 1;
	int OutputSheetRow = 0;
	int i = 0, j = 0;

	/*下面开始写入Output Excel*/
	OutExcelOperate.InitExcel();
	OutExcelOperate.OpenExcelFile(this->OutputExcelPath);
	OutExcelOperate.LoadSheet(CurSheet,1);
	OutputSheetRow = OutExcelOperate.GetRowCount();
	OutputSheetRow = 1;

	CString ReadCString;
	CString OutputCString;

	for(i = 0;i < this->InputExcelNum; i++) {
#if 0
		gDataCheck.ExcelCalcCStringTotalMoney = OutExcelOperate.GetCellString(OutputSheetRow + i + 1, 12);
		gDataCheck.ExcelCalcCStringRemainMoney = OutExcelOperate.GetCellString(OutputSheetRow + i + 1, 13);
		gDataCheck.InsertCStringTotalMoney = OutExcelOperate.GetCellString(OutputSheetRow + i + 1, 16);
		gDataCheck.InsertCStringRemainMoney = OutExcelOperate.GetCellString(OutputSheetRow + i + 1, 17);
		if(gDataCheck.ExcelCalcCStringTotalMoney != gDataCheck.InsertCStringTotalMoney) {
			OutExcelOperate.SetCellString(OutputSheetRow + i + 1, 18, CString("1"));
		}

		if(gDataCheck.ExcelCalcCStringRemainMoney != gDataCheck.InsertCStringRemainMoney) {
			OutExcelOperate.SetCellString(OutputSheetRow + i + 1, 19, CString("1"));
		}
#else
		double diffTotal = 0.0;
		double diffRemain = 0.0;

		gDoubleDataCheck.ExcelCalcCStringTotalMoney = OutExcelOperate.GetCellDouble(OutputSheetRow + i + 1, 12);
		gDoubleDataCheck.ExcelCalcCStringRemainMoney = OutExcelOperate.GetCellDouble(OutputSheetRow + i + 1, 13);
		gDoubleDataCheck.InsertCStringTotalMoney = OutExcelOperate.GetCellDouble(OutputSheetRow + i + 1, 16);
		gDoubleDataCheck.InsertCStringRemainMoney = OutExcelOperate.GetCellDouble(OutputSheetRow + i + 1, 17);

		diffTotal = gDoubleDataCheck.ExcelCalcCStringTotalMoney > gDoubleDataCheck.InsertCStringTotalMoney?\
			gDoubleDataCheck.ExcelCalcCStringTotalMoney - gDoubleDataCheck.InsertCStringTotalMoney:\
			gDoubleDataCheck.InsertCStringTotalMoney - gDoubleDataCheck.ExcelCalcCStringTotalMoney;

		diffRemain = gDoubleDataCheck.ExcelCalcCStringRemainMoney > gDoubleDataCheck.InsertCStringRemainMoney?\
			gDoubleDataCheck.ExcelCalcCStringRemainMoney - gDoubleDataCheck.InsertCStringRemainMoney:\
			gDoubleDataCheck.InsertCStringRemainMoney - gDoubleDataCheck.ExcelCalcCStringRemainMoney;

		if(diffTotal >= 0.01 || diffRemain >= 0.01) {
			OutExcelOperate.SetCellString(OutputSheetRow + i + 1, 18, CString("1"));
		} else {
			OutExcelOperate.SetCellString(OutputSheetRow + i + 1, 18, CString("0"));
		}
#endif
	}
	/*清理中间数据*/
#if 1
	/*先读取数据，再写入数据，替换掉公式算出来的结果*/
	{
		CString TmpCstring;
		unsigned int RowNum = this->InputExcelNum;
		unsigned int PrintRouNum = RowNum / 10;

		if(PrintRouNum == 0) PrintRouNum = 1;
		for(int i = 2; i <= RowNum+1; i++) {
			for(int j = 1;j <= 11; j++) {
				TmpCstring = OutExcelOperate.GetCellString(i,j);

				OutExcelOperate.SetCellString(i, j, TmpCstring);
			}

			if(i % PrintRouNum == 0)
				this->Printf("正在剔除公式(%u/%u)\r\n",i - 1,RowNum);
		}
	}

	

	{
		unsigned int RowNum = this->InputExcelNum;
		unsigned int PrintRouNum = RowNum / 10;
		if(PrintRouNum == 0) PrintRouNum = 1;
		for(int i = 1; i <= RowNum; i++) {
			for(int j = 19;j <= 28; j++) {
				OutExcelOperate.SetCellString(i, j, CString(""));
			}

			if(i % PrintRouNum == 0)
				this->Printf("正在剔除临时数据(%u/%u)\r\n",i,RowNum);
		}


	}

#endif
	OutExcelOperate.SaveasXSLFile(this->OutputExcel_DataCheckPath);
	OutExcelOperate.CloseExcelFile();
	OutExcelOperate.ReleaseExcel();

	CFileStatus Fstatus;
	if(CFile::GetStatus(this->OutputExcelPath,Fstatus,NULL)){
		CFile::Remove(this->OutputExcelPath);
	}
}

void CExcelHandler::Excel_AllHandler(void)
{
	OutputExcel_ModlePath = CExcelHandler::GetModulePath();
	OutputExcelPath = CExcelHandler::GetModulePath();
	OutputExcel_DataCheckPath = CExcelHandler::GetModulePath();

	OutputExcel_ModlePath += CString("\\..\\Input\\总表.xlsx");
	OutputExcelPath += CString("\\..\\Output\\output.xlsx");
	OutputExcel_DataCheckPath += CString("\\..\\Output\\汇总表.xlsx");

	CExcelHandler::RemoveOutputFile();

	CExcelHandler::FindInputFile();

	CExcelHandler::HandlerInputFile();	//处理Excel数据

	CExcelHandler::CreateOutputFile();

	CExcelHandler::OutputDataCheck();

	AfxMessageBox(_T("完成"));
}