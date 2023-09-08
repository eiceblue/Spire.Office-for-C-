#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"ToSVG.pdf";
	wstring outputFile = OUTPUTPATH"ToSVG/ToSVG.svg";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->SaveToFile(outputFile.c_str(), FileFormat::SVG);
	doc->Close();
}
