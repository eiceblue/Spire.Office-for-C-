#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"ConvertToOFD.ofd";
	wstring inputFile = DATAPATH"ConvertToOFD.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->SaveToFile(outputFile.c_str(), FileFormat::OFD);
}