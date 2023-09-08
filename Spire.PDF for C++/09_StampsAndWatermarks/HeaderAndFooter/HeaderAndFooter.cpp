#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"HeaderAndFooter.pdf";
	wstring inputFile1 = DATAPATH"Header.png";
	wstring inputFile2 = DATAPATH"Footer.png";
	wstring inputFile = DATAPATH"HeaderAndFooter.pdf";
	//Load the Pdf from disk
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	intrusive_ptr<PdfBrush> brush = PdfBrushes::GetBlack();
	intrusive_ptr<PdfPen> pen = new PdfPen(brush, 0.75f);
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 10.f, PdfFontStyle::Italic, true);

	intrusive_ptr<PdfStringFormat> rightAlign = new PdfStringFormat(PdfTextAlignment::Right);
	intrusive_ptr<PdfStringFormat> leftAlign = new PdfStringFormat(PdfTextAlignment::Left);
	rightAlign->SetMeasureTrailingSpaces(true);
	rightAlign->SetMeasureTrailingSpaces(true);
	intrusive_ptr<PdfMargins> margin = doc->GetPageSettings()->GetMargins();

	float space = font->GetHeight() * 0.75f;
	float x = 0;
	float y = 0;
	float width = 0;

	//Create a new pdf document
	intrusive_ptr<PdfDocument> newPdf = new PdfDocument();

	for (int i = 0; i < doc->GetPages()->GetCount(); i++)

	{
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(i);
		//Add new page

		intrusive_ptr<PdfPageBase>  newPage = newPdf->GetPages()->Add(page->GetSize(), new PdfMargins(0.0f));

		newPage->GetCanvas()->SetTransparency(0.5f);
		x = margin->GetLeft();
		width = page->GetCanvas()->GetClientSize()->GetWidth() - margin->GetLeft() - margin->GetRight();
		y = margin->GetTop() - space;

		//Draw header line
		newPage->GetCanvas()->DrawLine(pen, x, y + 15, x + width, y + 15);
		y = y + 10 - font->GetHeight();

		//Draw header image into newPage
		newPage->GetCanvas()->SetTransparency(0.5f);

		intrusive_ptr<PdfImage> headerImage = PdfImage::FromFile(inputFile1.c_str());
		newPage->GetCanvas()->DrawImage(headerImage, new PointF(0, 0));

		//Draw header text into newPage
		newPage->GetCanvas()->DrawString(L"Demo of Spire.Pdf", font, brush, x + width, y, rightAlign);

		//Draw footer image into newPage
		intrusive_ptr<PdfImage> footerImage = PdfImage::FromFile(inputFile2.c_str());
		newPage->GetCanvas()->DrawImage(footerImage, new PointF(0, newPage->GetCanvas()->GetClientSize()->GetHeight() - footerImage->GetPhysicalDimension()->GetHeight()));

		brush = PdfBrushes::GetDarkBlue();
		font = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Bold, true);

		y = newPage->GetCanvas()->GetClientSize()->GetHeight() - margin->GetBottom() - font->GetHeight();

		//Draw footer text into newPage
		newPage->GetCanvas()->DrawString(L"Created by E-iceblue Co,.Ltd", font, brush, x, y, leftAlign);

		newPage->GetCanvas()->SetTransparency(1);

		//Draw the page into newPage

		Object::Convert<PdfGraphicsWidget>(page->CreateTemplate())->Draw(newPage->GetCanvas(), new PointF(0, 0));
	}

	//Save the document
	newPdf->SaveToFile(outputFile.c_str());
	newPdf->Close();
}