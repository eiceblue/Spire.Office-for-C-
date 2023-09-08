#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DrawShapeWithGradientFill.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	intrusive_ptr<PdfLinearGradientBrush> brush1 = new PdfLinearGradientBrush(new RectangleF(new PointF(100, 100), new SizeF(200, 120)),
		new PdfRGBColor(Spire::Common::Color::GetBlueViolet()), new PdfRGBColor(Spire::Common::Color::GetDarkBlue()), PdfLinearGradientMode::Horizontal);
	page->GetCanvas()->DrawRectangle(brush1, new RectangleF(new PointF(100, 100), new SizeF(200, 120)));
	intrusive_ptr<PdfRadialGradientBrush> brush2 = new PdfRadialGradientBrush(new PointF(200.f, 400.f), 100.f, new PointF(300.f, 500.f),
		100.f, new PdfRGBColor(Spire::Common::Color::GetSkyBlue()), new PdfRGBColor(Spire::Common::Color::GetDarkBlue()));
	page->GetCanvas()->DrawEllipse(brush2, new RectangleF(new PointF(100, 300), new SizeF(200, 200)));
	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();

}


