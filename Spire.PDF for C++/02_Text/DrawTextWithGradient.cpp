#include "../pch.h"

using namespace Spire::Pdf;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path +L"DrawTextWithGradient.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	intrusive_ptr<RectangleF> rect = new RectangleF(new PointF(0, 0), new SizeF(300.f, 100.f));
	intrusive_ptr<PdfLinearGradientBrush> brush = new PdfLinearGradientBrush(rect, new PdfRGBColor(Color::GetRed()), new PdfRGBColor(Color::GetBlue()),
		PdfLinearGradientMode::Horizontal);
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 20.f, PdfFontStyle::Underline, true);

	page->GetCanvas()->DrawString(L"Welcome to E-iceblue!", font, brush, new PointF(1, 100));

	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

