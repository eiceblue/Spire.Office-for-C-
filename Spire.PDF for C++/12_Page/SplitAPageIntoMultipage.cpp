#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"SplitAPageIntoMultipage.pdf";
	wstring inputFile = DATAPATH"PDFTemplate_N.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first page
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	//Create a new Pdf
	intrusive_ptr<PdfDocument> newPdf = new PdfDocument();

	//Remove all the margins
	newPdf->GetPageSettings()->GetMargins()->SetAll(0);

	//Set the page size of new Pdf
	newPdf->GetPageSettings()->SetWidth(page->GetSize()->GetWidth());
	newPdf->GetPageSettings()->SetHeight(page->GetSize()->GetHeight() / 2);

	//Add a new page
	intrusive_ptr<PdfPageBase> newPage = newPdf->GetPages()->Add();

	intrusive_ptr<PdfTextLayout> format = new PdfTextLayout();
	format->SetBreak(PdfLayoutBreakType::FitPage);
	format->SetLayout(PdfLayoutType::Paginate);

	//Draw the page in the new page
	page->CreateTemplate()->Draw(newPage, new PointF(0, 0), format);

	//Save the document
	newPdf->SaveToFile(outputFile.c_str());
	newPdf->Close();
}
