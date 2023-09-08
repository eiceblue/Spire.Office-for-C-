#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"GroupRowsAndColumns.xls";
	wstring outputFile = outputFolder + L"GroupRowsAndColumns_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Grouping rows
	sheet->GroupByRows(1, 5, false);
	//Grouping columns
	sheet->GroupByColumns(1, 3, false);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
