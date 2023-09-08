#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring inputFile = DATAPATH"TextWaterMark.pdf";
	wstring outputFile = OUTPUTPATH"TextWaterMark.pdf";

	//Create a pdf document and load file from disk
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first page
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	//Draw text watermark
	intrusive_ptr<PdfTilingBrush> brush = new PdfTilingBrush(new SizeF(page->GetCanvas()->GetClientSize()->GetWidth() / 2, page->GetCanvas()->GetClientSize()->GetHeight() / 3));
	brush->GetGraphics()->SetTransparency(0.3f);
	brush->GetGraphics()->Save();
	brush->GetGraphics()->TranslateTransform(brush->GetSize()->GetWidth() / 2, brush->GetSize()->GetHeight() / 2);
	brush->GetGraphics()->RotateTransform(-45);

	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 24);
	intrusive_ptr<PdfStringFormat> stringformat = new PdfStringFormat(PdfTextAlignment::Center);
	brush->GetGraphics()->DrawString(L"Spire.Pdf Demo", font, PdfBrushes::GetViolet(), 0, 0, stringformat);
	brush->GetGraphics()->Restore();
	brush->GetGraphics()->SetTransparency(1);
	intrusive_ptr<PointF> point = new PointF(0, 0);
	intrusive_ptr<RectangleF> rec = new RectangleF(point, page->GetCanvas()->GetClientSize());
	page->GetCanvas()->DrawRectangle(brush, rec);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
