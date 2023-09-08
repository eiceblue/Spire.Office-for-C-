#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"InsertEmptyPage.pdf";
	wstring inputFile = DATAPATH"Sample.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//insert a blank page as the second page
	doc->GetPages()->Insert(1);
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
