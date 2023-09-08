#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"InvisibleAllPdfLayers.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	wstring inputFile = DATAPATH"Template_Pdf_5.pdf";
	doc->LoadFromFile(inputFile.c_str());

	for (int i = 0; i < doc->GetLayers()->GetCount(); i++)
	{
		//Set all the Pdf layers invisible.
		doc->GetLayers()->GetItem(i)->SetVisibility(PdfVisibility::Off);
	}

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
