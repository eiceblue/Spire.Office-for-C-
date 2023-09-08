#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"InsertEmptyPageAtEnd.pdf";
	wstring inputFile = DATAPATH"MultipagePDF.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	
	intrusive_ptr<PdfMargins> tempVar = new PdfMargins(0, 0);
	doc->GetPages()->Add(new SizeF(PdfPageSize::A4()), tempVar);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
