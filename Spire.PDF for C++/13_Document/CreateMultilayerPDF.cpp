#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateMultilayerPDF.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	// Creates a page
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();

	//Create text
	LPCWSTR_S text = L"Welcome to evaluate Spire.PDF for .NET !";

	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat(PdfTextAlignment::Left);

	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Calibri", 15.0f, PdfFontStyle::Regular, true);
	intrusive_ptr<PdfBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetBlack()));

	float x = 50;
	float y = 50;
	intrusive_ptr<PointF> pf = new PointF(x, y);
	// Draw text layer
	page->GetCanvas()->DrawString(text, font, brush, pf, format);

	intrusive_ptr<SizeF> size = font->MeasureString(L"Welcome to  evaluate", format);

	intrusive_ptr<SizeF> size2 = font->MeasureString(L"Spire.PDF for .NET", format);

	// Loads an image 

	wstring inputFile = DATAPATH"MultilayerImage.png";
	intrusive_ptr<PdfImage> image = PdfImage::FromFile(inputFile.c_str());
	pf = new PointF(x + size->GetWidth(), y);
	// Draw image layer
	page->GetCanvas()->DrawImage(image, pf, size2);
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}