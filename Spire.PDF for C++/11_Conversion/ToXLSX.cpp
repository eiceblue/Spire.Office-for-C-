#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"ChartSample.pdf";
	wstring outputFile = OUTPUTPATH"ToXLSX.xlsx";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->SaveToFile(outputFile.c_str(), FileFormat::XLSX);
}
