#ifndef FILEHANDLER_H
#define FILEHANDLER_H
#pragma once
// excelcompare.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"
#include <iostream>
#include <string>  
#include <cmath>
#include "atlconv.h"
#include "stdlib.h"
#include <fstream>

#import "C:\\Program Files (x86)\\Common Files\\microsoft shared\\OFFICE16\\mso.dll" rename("RGB", "MsoRGB")
#import "C:\\Program Files (x86)\\Common Files\\microsoft shared\\VBA\\VBA6\\VBE6EXT.OLB"
#import "C:\\Program Files (x86)\\Microsoft Office\\Office16\\excel.exe"\
	 rename("DialogBox", "ExcelDialogBox") \
	 rename("RGB", "ExcelRGB") \
	 rename("CopyFile", "ExcelCopyFile") \
	 rename("ReplaceText", "ExcelReplaceText") \
	 exclude("IFont", "IPicture") no_dual_interfaces
#define BUFFER_SIZE 100
using namespace std;

class Openfiles
{
public:
	Openfiles();
	Excel::RangePtr pRange;
	Excel::RangePtr pRange2;
	std::fstream myfile;
	// Create Excel Application Object pointer  
	Excel::_ApplicationPtr pXL;
	Excel::_ApplicationPtr pXL2;
	int fileopen(void);
};

Openfiles::Openfiles(void)
{
	cout << "Object is being created" << endl;
}
//member functions
int Openfiles::fileopen()
{

	HRESULT hr = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hr))
	{
		std::cout << "Failed to initialize COM library. Error code = 0x"
			<< std::hex << hr << std::endl;
		return hr;
	}



	if (FAILED(pXL.CreateInstance("Excel.Application")))
	{
		std::cout << "Failed to initialize Excel::_Application!" << std::endl;
		return 0;
	}
	if (FAILED(pXL2.CreateInstance("Excel.Application")))
	{
		std::cout << "Failed to initialize Excel::_Application!" << std::endl;
		return 0;
	}
	// Open the Excel Workbook, but don't make it visible  
	pXL->Workbooks->Open(L"C:\\Users\\1145660\\Desktop\\TEMP\\excelcompare\\excelcompare_KSU\\excel1.xlsx");
	// Open the Excel Workbook, but don't make it visible  
	pXL2->Workbooks->Open(L"C:\\Users\\1145660\\Desktop\\TEMP\\excelcompare\\excelcompare_KSU\\excel2.xlsx");

	// Access Excel Worksheet and return pointer to Worksheet cells  
	Excel::_WorksheetPtr pWksheet = pXL->ActiveSheet;
	// Access Excel Worksheet and return pointer to Worksheet cells  
	Excel::_WorksheetPtr pWksheet2 = pXL2->ActiveSheet;

	pWksheet->Name = L"Sheet1";
	pWksheet2->Name = L"Sheet1";

	pRange = pWksheet->Cells;
	pRange2 = pWksheet2->Cells;

	
	myfile.open("C:\\Users\\1145660\\Desktop\\TEMP\\excelcompare\\excelcompare_KSU\\log.txt", std::fstream::out);
}

#endif