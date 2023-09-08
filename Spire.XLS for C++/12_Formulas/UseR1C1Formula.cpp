#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UseR1C1Formula.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetNumberValue(1);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetNumberValue(2);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetNumberValue(3);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetNumberValue(4);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(5);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(6);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetNumberValue(7);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetNumberValue(8);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetNumberValue(9);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetText(L"Sum:");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Right);
	//Write array  R1C1 formula
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetFormulaR1C1(L"=SUM(R[-3]C[-2]:R[-1]C)");

	//Calculate Formulas
	workbook->CalculateAllValue();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}