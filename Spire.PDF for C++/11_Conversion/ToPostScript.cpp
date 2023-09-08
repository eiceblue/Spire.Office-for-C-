#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"ToPostScript.pdf";
	wstring outputFile = OUTPUTPATH"ToPostScript.ps";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->SaveToFile(outputFile.c_str(), FileFormat::POSTSCRIPT);
	doc->Close();
}

