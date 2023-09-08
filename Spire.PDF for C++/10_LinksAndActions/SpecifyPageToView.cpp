#include "pch.h"

using namespace Spire::Common;
using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"Sample.pdf";
	wstring outputFile = OUTPUTPATH"SpecifyPageToView.pdf";
	
	//Create a pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	//Load an existing pdf from disk
	doc->LoadFromFile(inputFile.c_str());

	//Create a PdfDestination with specific page, location and 50% zoom factor
	intrusive_ptr<PdfDestination> dest = new PdfDestination(2, new PointF(0, 100), 0.5f);

	//Create GoToAction with dest
	intrusive_ptr<PdfGoToAction> action = new PdfGoToAction(dest);

	//Set open action
	doc->SetAfterOpenAction(action);
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
