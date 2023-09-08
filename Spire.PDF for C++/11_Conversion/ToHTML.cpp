#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"ToHTML.pdf";
	wstring outputFile = OUTPUTPATH"ToHTML.html";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->SaveToFile(outputFile.c_str(), FileFormat::HTML);
	doc->Close();
}
