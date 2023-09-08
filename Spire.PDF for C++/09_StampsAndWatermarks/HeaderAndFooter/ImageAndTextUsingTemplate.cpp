#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring inputFile = DATAPATH"PDFTemplate_HF.pdf";
	wstring outputFile = OUTPUTPATH"ImageAndTextUsingTemplate.pdf";

	//Load the Pdf from disk
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	//Get the first page
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	//Get the margins of Pdf
	intrusive_ptr<PdfMargins> margin = doc->GetPageSettings()->GetMargins();

	//Define font and brush
	LPCWSTR_S impact = L"Impact";
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(impact, 14.f, PdfFontStyle::Regular);
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetGray()));

	//Load an image
	wstring inputFile1 = DATAPATH"E-iceblueLogo.png";
	intrusive_ptr<PdfImage> image = PdfImage::FromFile(inputFile1.c_str());

	//Specify the image size
	intrusive_ptr<SizeF> imageSize = new SizeF(image->GetWidth() / 2, image->GetHeight() / 2);

	//Create a header template
	intrusive_ptr<PdfTemplate> headerTemplate = new PdfTemplate(page->GetActualSize()->GetWidth() - margin->GetLeft() - margin->GetRight(), imageSize->GetHeight());

	//Draw the image in the template
	headerTemplate->GetGraphics()->DrawImage(image, new PointF(0, 0), imageSize);

	//Create a retangle
	intrusive_ptr<RectangleF> rect = headerTemplate->GetBounds();

	//string format
	intrusive_ptr<PdfStringFormat> format1 = new PdfStringFormat(PdfTextAlignment::Right, PdfVerticalAlignment::Middle);

	//Draw a string in the template
	headerTemplate->GetGraphics()->DrawString(L"Header", font, brush, rect, format1);

	//Create a footer template and draw a text
	intrusive_ptr<PdfTemplate> footerTemplate = new PdfTemplate(page->GetActualSize()->GetWidth() - margin->GetLeft() - margin->GetRight(), imageSize->GetHeight());
	intrusive_ptr<PdfStringFormat> format2 = new PdfStringFormat(PdfTextAlignment::Center, PdfVerticalAlignment::Middle);
	footerTemplate->GetGraphics()->DrawString(L"Footer", font, brush, rect, format2);

	float x = margin->GetLeft();
	float y = 0;

	//Draw the header template on page at specified location
	page->GetCanvas()->DrawTemplate(headerTemplate, new PointF(x, y));

	//Draw the footer template on page at specified location
	y = page->GetActualSize()->GetHeight() - footerTemplate->GetHeight() - 10;
	page->GetCanvas()->DrawTemplate(footerTemplate, new PointF(x, y));

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

