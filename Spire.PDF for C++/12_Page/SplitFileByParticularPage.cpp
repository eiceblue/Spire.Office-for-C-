#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"SplitFileByParticularPage.pdf";
	wstring inputFile = DATAPATH"Sample.pdf";

	intrusive_ptr<PdfDocument> oldPdf = new PdfDocument();
	oldPdf->LoadFromFile(inputFile.c_str());

	//Create a new PDF document
	intrusive_ptr<PdfDocument> newPdf = new PdfDocument();

	//Initialize a new instance of PdfPageBase class
	intrusive_ptr<PdfPageBase> page;

	//Specify the pages which you want them to be split
	for (int i = 1; i < 3; i++)
	{
		intrusive_ptr<PdfMargins> tempVar = new PdfMargins(0);
		page = newPdf->GetPages()->Add(oldPdf->GetPages()->GetItem(i)->GetSize(), tempVar);

		//Create template of the oldPdf page and draw into newPdf page
		oldPdf->GetPages()->GetItem(i)->CreateTemplate()->Draw(page, new PointF(0, 0));
	}

	//Save the document
	newPdf->SaveToFile(outputFile.c_str());
	newPdf->Close();
}
