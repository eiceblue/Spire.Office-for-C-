#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring inputFile = DATAPATH"PDFTemplate-Az.pdf";
	wstring outputFile = OUTPUTPATH"RemovePageMargins.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	// Creates a new page
	intrusive_ptr<PdfDocument> newDoc = new PdfDocument();

	// Get page margins of source pdf page
	intrusive_ptr<PdfMargins> margins = doc->GetPageSettings()->GetMargins();
	float top = margins->GetLeft();
	float bottom = margins->GetBottom();
	float left = margins->GetLeft();
	float right = margins->GetRight();

	for (int i = 0; i < doc->GetPages()->GetCount(); i++)
	{
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

		intrusive_ptr<PdfMargins> tempVar = new PdfMargins(0);
		intrusive_ptr<PdfPageBase> newPage = newDoc->GetPages()->Add(new SizeF(page->GetSize()->GetWidth() - left - right, page->GetSize()->GetHeight() - top - bottom), tempVar);

		// Draws the content of the source page onto the new document page
		newPage->GetCanvas()->DrawTemplate(page->CreateTemplate(), new PointF(-left, -top));
	}
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
