#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"ProtectedWorkbook.xlsx";
	wstring outputFile = output + L"DetectProtection.txt";
	wfstream ofs;

	bool value = Workbook::IsPasswordProtected(inputFile.c_str());
	wstring* boolvalue = new wstring();
	if (value)
	{
		boolvalue->append(L"Yes");
	}
	else
	{
		boolvalue->append(L"No");
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *boolvalue << endl;
	ofs.close();
}