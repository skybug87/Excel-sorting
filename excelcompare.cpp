// excelcompare.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"
#include <iostream>
#include <string>  
#include <cmath>
#include "atlconv.h"
#include "stdlib.h"
#include <fstream>
#include "Filehandler.h"


bool compare_float(float x, float y, float epsilon = 0.01f)
{
	if (fabs(x - y) < epsilon)
		return true; //they are same
	return false; //they are not same
}

bool IsUnicodeDigit(_bstr_t ch)
{
	WORD type;
	return GetStringTypeW(CT_CTYPE1, ch, 1, &type) &&
		(type & C1_DIGIT);
}
int main()
{

	Openfiles doc;
	int count = 0;
	doc.fileopen();
	//Read excel Values. Read an Excel data cell. (Note Excel cells start from index = 1)
	//capture ID and get limits
	for (int i = 2; i < 5; i++)
	{

		//Get cell ID
		bool pfound = false;
		bool limitmatched = false;
		doc.pRange->Item[i][2];
		_bstr_t pcodes = _bstr_t(doc.pRange->Item[i][1]);
		char pcode[100];
		strcpy_s(pcode, (char*)pcodes);
		
		std::cout << "Searching for ID: " << pcode << std::endl;
		//check limits
		bool isdigit;
		//add for viewing pass/fail
		//capture as a string
		_bstr_t  limitlower = _bstr_t(doc.pRange->Item[i][2]);
		_bstr_t limitupper = _bstr_t(doc.pRange->Item[i][3]);

		//check if string is contains digits
		isdigit = IsUnicodeDigit(limitlower);
		char lower[100];
		char upper[100];
		double limitlowerd = 0;
		double limitupperd = 0;
		double limitlowerd2 = 0;
		double limitupperd2 = 0;
		strcpy_s(upper, (char*)limitupper);
		strcpy_s(lower, (char*)limitlower);

		//if limit is not a digit
		if ((!isdigit) || ((strstr(upper, "0x")) || (strstr(lower, "0x"))))
		{
			std::cout << "Searching for lower: " << lower << std::endl;
			std::cout << "Searching for upper: " << upper << std::endl;
		}
		else
		{
			limitlowerd = doc.pRange->Item[i][2];
			limitupperd = doc.pRange->Item[i][3];
			std::cout << "Searching for limitlowerd: " << limitlowerd << std::endl;
		}



		for (int j = 2; j < 5; j++)
		{
			//get ID out of 2nd excel doc
			_bstr_t  pcodessearch = _bstr_t(doc.pRange2->Item[j][1]);
			char pcodesearch[100];
			strcpy_s(pcodesearch, (char*)pcodessearch);
			//str2 = "0x" + str2;
			//if pcode matches grab upper and lower limits out of tear
			std::cout << "Searching for Pcodesearch match: " << pcodesearch << std::endl;
			if (strstr(pcodesearch, pcode))
			{
				pfound = true;
				_bstr_t upperlimit2 = _bstr_t(doc.pRange2->Item[j][3]);
				_bstr_t lowerlimit2 = _bstr_t(doc.pRange2->Item[j][2]);
				char lower2[100];
				char upper2[100];

				strcpy_s(upper2, (char*)upperlimit2);
				strcpy_s(lower2, (char*)lowerlimit2);

				//if limit is not a digit
				if ((!isdigit) || ((strstr(upper2, "xx")) || (strstr(upper2, "0x")) || (strstr(lower2, "0x"))))
				{
					std::cout << "Searching for lower: " << lower << std::endl;
					std::cout << "Searching for upper: " << upper << std::endl;
					if (strstr(upper2, upper) || strstr(lower, lower2))
					{
						limitmatched = true;
					}

				}
				else
				{
					limitlowerd2 = doc.pRange2->Item[j][3];
					limitupperd2 = doc.pRange2->Item[j][2];
					std::cout << "Searching for limitlowerd: " << limitlowerd << std::endl;

					switch ((compare_float(limitupperd2, limitupperd)) || (compare_float(limitlowerd2, limitlowerd)))
					{
					case true:
						limitmatched = true;
						std::cout << "They are equivalent" << std::endl;
						break;
					case false:
						std::cout << "They are not equivalent" << std::endl;
						break;
					}
				}
			}

		}

		if (!pfound && (pcode[0] != NULL)) {
			doc.myfile << pcode << " not found " << std::endl;
			count++;
		}
		if (!limitmatched && pfound && pcode != NULL) {
			doc.myfile << pcode << " upper limits not matched. New upper limit " << limitupperd2 << " Old upper limit " << limitupperd << std::endl;
			doc.myfile << pcode << " lower limits not matched. New lower limit " << limitlowerd2 << " Old lower limit " << limitlowerd << std::endl;
		}
	}
	doc.myfile << count << " total not found " << std::endl;
	doc.myfile.close();
	
	// Switch off alert prompting to save as 
	doc.pXL->PutDisplayAlerts(false);

	// And switch back on again...
	doc.pXL->PutDisplayAlerts(true);

	doc.pXL->Quit();
	
	std::cout << "Hello World!\n"; 
}

