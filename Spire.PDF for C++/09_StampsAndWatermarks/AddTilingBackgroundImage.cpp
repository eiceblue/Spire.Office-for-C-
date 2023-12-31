#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"AddTilingBackgroundImage.pdf";
	wstring inputFile = DATAPATH"PDFTemplate_N.pdf";

	//Load the Pdf from disk
	intrusive_ptr<PdfDocument> document = new PdfDocument();
	document->LoadFromFile(inputFile.c_str());
	wstring inputFile_Logo = DATAPATH"E-iceblueLogo.png";
	//Load an image
	intrusive_ptr<PdfImage> image = PdfImage::FromFile(inputFile_Logo.c_str());

	for (int i = 0; i < document->GetPages()->GetCount(); i++)
	{
		intrusive_ptr<PdfPageBase> page = document->GetPages()->GetItem(i);
		//Create PdfTilingBrush
		intrusive_ptr<PdfTilingBrush> brush = new PdfTilingBrush(new SizeF(page->GetCanvas()->GetSize()->GetWidth() / 3.0f, page->GetCanvas()->GetSize()->GetHeight() / 5.0f));

		//Set the transparency
		brush->GetGraphics()->SetTransparency(0.3f);

		//Draw image on brush graphics
		brush->GetGraphics()->DrawImage(image, new PointF((brush->GetSize()->GetWidth() - image->GetWidth()) / 2.f, (brush->GetSize()->GetHeight() - image->GetHeight()) / 2.f));

		//use the brush to draw rectangle
		page->GetCanvas()->DrawRectangle(brush, new RectangleF(new PointF(0, 0), page->GetCanvas()->GetSize()));
	}

	//Save the document
	document->SaveToFile(outputFile.c_str());
	document->Close();
}
