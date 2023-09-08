#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"ZoomToPageContents.pdf";
	wstring inputFile = DATAPATH"PDFTemplate_N.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Create a newDoc
	intrusive_ptr<PdfDocument> newDoc = new PdfDocument();

	for (int i = 0; i < doc->GetPages()->GetCount(); i++)
	{
		intrusive_ptr<PdfMargins> tempVar = new PdfMargins(0, 0);
		intrusive_ptr<PdfPageBase> newPage = newDoc->GetPages()->Add(new SizeF(PdfPageSize::A3()), tempVar);
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(i);
		//Zoom content to the new page
		newPage->GetCanvas()->ScaleTransform(newPage->GetActualSize()->GetWidth() / page->GetActualSize()->GetWidth(), (newPage->GetActualSize()->GetHeight() / page->GetActualSize()->GetHeight()));

		//Draw the page to new page
		newPage->GetCanvas()->DrawTemplate(page->CreateTemplate(), new PointF(0, 0));
	}

	///Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
