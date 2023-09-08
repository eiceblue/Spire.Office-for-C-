#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"SVGToPDF.pdf";
	wstring inputFile = DATAPATH"template.svg";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromSvg(inputFile.c_str());
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}