#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"ToDoc.pdf";
	wstring outputFile = OUTPUTPATH"ToDoc.doc";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->SaveToFile(outputFile.c_str(), FileFormat::DOC);
	doc->Close();
}