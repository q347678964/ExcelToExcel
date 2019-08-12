#ifndef EXCEL_HANDLER
#define EXCEL_HANDLER

#include "FormatChange.h"
#include "afxwin.h"
#include "UtilityFunc.h"

#define MAX_INTPUT_EXCEL 1000
#define MAX_CELL	19

typedef CString (*InputCStringHandler)(CString);

typedef struct _IN_OUT_CONFIG_S_ {
	unsigned int InputRow;
	unsigned int InputColumn;
	CString CurPointCString;
	unsigned int OutputColumn;
	InputCStringHandler Hld;
}IN_OUT_CONFIG_S;

class CExcelHandler:public FormatChange, public UtilityFunc
{
	public:
		IN_OUT_CONFIG_S gInOutConfig[MAX_INTPUT_EXCEL][MAX_CELL];
		CString  InputExcelPath[MAX_INTPUT_EXCEL];
		unsigned int InputExcelNum;
		CString OutputExcel_ModlePath;
		CString OutputExcelPath;
		CString OutputExcel_DataCheckPath;
	public:
		CExcelHandler(void);
		void DebugUpdate(void);
		void RemoveOutputFile(void);
		void FindInputFile(void);
		void HandlerInputFile(void);
		void CreateOutputFile(void);
		void OutputDataCheck(void);
		void Excel_AllHandler(void);
		CString DebugInfoString;
};

#endif
