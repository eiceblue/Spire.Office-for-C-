#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

void PageLable_2()
{

	wstring outputFile = OUTPUTPATH"PageLable_2.pdf";
	wstring inputFile = DATAPATH"hasLable.pdf";

	intrusive_ptr<PdfDocument> newdoc = new PdfDocument();
	newdoc->LoadFromFile(inputFile.c_str());

	intrusive_ptr<PdfPageLabels> pageLabels = newdoc->GetPageLabels();
	pageLabels->AddRange(0, PdfPageLabels::Decimal_Arabic_Numerals_Style(), L"new lable");
	newdoc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	newdoc->Dispose();
}

int main()
{
	wstring outputFile = OUTPUTPATH"PageLable_1.pdf";
	wstring inputFile = DATAPATH"Sample.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	doc->SetPageLabels(new PdfPageLabels());
	doc->GetPageLabels()->AddRange(0, PdfPageLabels::Decimal_Arabic_Numerals_Style(), L"label test ");
	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);

	PageLable_2();
}
