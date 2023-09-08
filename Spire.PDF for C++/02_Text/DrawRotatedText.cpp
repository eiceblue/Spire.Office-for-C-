#include "../pch.h"

using namespace Spire::Pdf;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path +L"DrawRotatedText.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
	intrusive_ptr<PdfRGBColor> color = new PdfRGBColor(Color::GetBlue());
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(color);
	LPCWSTR_S text = L"This is a text";
	page->GetCanvas()->DrawString(text, font, brush, 20, 30);
	page->GetCanvas()->DrawString(text, font, brush, 20, 150);
	intrusive_ptr<PdfGraphicsState> state = page->GetCanvas()->Save();
	intrusive_ptr<PointF> point1 = new PointF(20, 30);
	page->GetCanvas()->RotateTransform(90, point1);
	page->GetCanvas()->DrawString(text, font, brush, point1);
	page->GetCanvas()->Restore(state);

	intrusive_ptr<PdfGraphicsState> state2 = page->GetCanvas()->Save();
	intrusive_ptr<PointF> point2 = new PointF(20, 150);
	page->GetCanvas()->RotateTransform(-90, point2);
	page->GetCanvas()->DrawString(text, font, brush, point2);
	page->GetCanvas()->Restore(state2);

	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

