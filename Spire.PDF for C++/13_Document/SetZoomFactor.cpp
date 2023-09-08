#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile =OUTPUTPATH"SetZoomFactor.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	wstring inputFile = DATAPATH"SetZoomFactor.pdf";
	doc->LoadFromFile(inputFile.c_str());

	//Get the first page
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	//Set pdf destination
	intrusive_ptr<PdfDestination> dest = new PdfDestination(page);
	dest->SetMode(PdfDestinationMode::Location);
	dest->SetLocation(new PointF(-40.0f, -40.0f));
	dest->SetZoom(0.6f);

	//Set action
	intrusive_ptr<PdfGoToAction> gotoAction = new PdfGoToAction(dest);
	doc->SetAfterOpenAction(gotoAction);
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
