#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"DeletePage.pdf";
	wstring inputFile = DATAPATH"DeletePage.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Delete the fifth page
	doc->GetPages()->RemoveAt(4);
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
