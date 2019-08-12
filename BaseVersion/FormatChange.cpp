#include "stdafx.h"
#include "resource.h"

#include "FormatChange.h"


FormatChange::FormatChange(){

}
FormatChange::~FormatChange(){

}
/*
*	CStringToChar-��CStringת��ΪChar*������.�����漰unicode->char
*	@Context	CString
*	@Char		ת��֮��Ĵ洢��ַ����Ҫ���ⲿ����,�����CStringһ����.
*	@Ret		ת��֮��Ĵ�С,
*/

int FormatChange::CStringToChar(CString Context,char *Char){
	int Len =WideCharToMultiByte(CP_ACP,0,Context.GetBuffer(Context.GetLength()),-1,NULL,0,NULL,NULL); //���㳤��
	WideCharToMultiByte(CP_ACP,0,Context.GetBuffer(Context.GetLength()),-1,Char,Len,NULL,NULL );	//ת��
	return Len;
}
/*
*	CharToCString-��charת��Ϊunicode��CString
*	@Context	Char��ַ
*	@Ret		CString
*/ 
CString FormatChange::CharToCString(char *Char){
	return CString(Char);
}

int FormatChange::WcharToChar(WCHAR *Wchar,char *Char){
	sprintf(Char,"%ws",Wchar);
	return 0;
}
/*
*	CStringToInt-��CStringתΪint
*	@Context	CString
*	@Ret		int
*/ 
int FormatChange::CStringToInt(CString Context){
	int ret;
	swscanf(Context,_T("%d"),&ret); 
	return ret;
}
/*
*	CStringToInt-��intתΪCString
*	@int	Num
*	@Ret	CString
*/ 
CString FormatChange::IntToCString(int Num){
	CString TempCString;
	TempCString.Format(_T("%d"),Num);
	return TempCString;
}
void FormatChange::CImage2CBitmap(CImage &Image,CBitmap &Bitmap){
	HBITMAP hbmp =(HBITMAP)Image.operator HBITMAP();		//CImage -> Hbitmap
	this->Hbitmap2CBitmap(hbmp,Bitmap);				//Hbitmap->Cbitmap
}

void FormatChange::Hbitmap2CBitmap(HBITMAP &Hbitmap,CBitmap &Bitmap){		//Hbitmap->Cbitmap
	Bitmap.DeleteObject();
	Bitmap.Attach(Hbitmap);
}

void FormatChange::CBitmap2Hbitmap(CBitmap &Bitmap,HBITMAP &Hbitmap){		//Cbitmap -> Hbitmap
	Hbitmap=(HBITMAP)Bitmap;
}



wchar_t* FormatChange::GB2312ToUnicode(const char* szGBString)
{
    UINT nCodePage = 936; //GB2312
 
    int nLength=MultiByteToWideChar(nCodePage,0,szGBString,-1,NULL,0);
 
    wchar_t* pBuffer = new wchar_t[nLength+1];
 
    MultiByteToWideChar(nCodePage,0,szGBString,-1,pBuffer,nLength);
 
    pBuffer[nLength]=0;
 
    return pBuffer;
}
 