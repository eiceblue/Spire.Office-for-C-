#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"ViewerPreference.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	wstring inputFile = DATAPATH"ViewerPreference.pdf";
	doc->LoadFromFile(inputFile.c_str());

	//Set view reference
	doc->GetViewerPreferences()->SetCenterWindow(true);
	doc->GetViewerPreferences()->SetDisplayTitle(false);
	doc->GetViewerPreferences()->SetFitWindow(false);
	doc->GetViewerPreferences()->SetHideMenubar(true);
	doc->GetViewerPreferences()->SetHideToolbar(true);
	doc->GetViewerPreferences()->SetPageLayout(PdfPageLayout::SinglePage);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
