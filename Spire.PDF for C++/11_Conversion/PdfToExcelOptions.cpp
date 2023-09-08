#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"PdfToExcelOptions.xlsx";
	wstring inputFile = DATAPATH"PdfToXlsxOptions.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->GetConvertOptions()->SetPdfToXlsxOptions(new XlsxLineLayoutOptions(false, true, true));
	doc->SaveToFile(outputFile.c_str(), FileFormat::XLSX);
	doc->Close();
}