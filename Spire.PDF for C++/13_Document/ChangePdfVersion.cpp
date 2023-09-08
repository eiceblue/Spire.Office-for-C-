#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"ChangePdfVersion.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	wstring inputFile = DATAPATH"ChangePdfVersion.pdf";
	doc->LoadFromFile(inputFile.c_str());

	//Change the pdf version
	doc->GetFileInfo()->SetVersion(PdfVersion::Version1_6);
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
