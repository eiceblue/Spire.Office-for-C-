#include "pch.h"

using namespace Spire::Common;
using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"OpenAction.pdf";
	wstring outputFile = OUTPUTPATH"RemoveOpenAction.pdf";
	
	//Create a pdf document
	intrusive_ptr<PdfDocument> document = new PdfDocument();

	//Load an existing pdf from disk
	document->LoadFromFile(inputFile.c_str());

	//Remove action
	document->SetAfterOpenAction();

	//Save the document
	document->SaveToFile(outputFile.c_str());
	document->Close();
}
