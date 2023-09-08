#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DrawingTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DrawRectangles.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfGraphicsState> state = page->GetCanvas()->Save();
	int x = 130;
	int y = 100;
	int width = 300;
	int height = 400;
	intrusive_ptr<PdfPen> pen = new PdfPen(new PdfRGBColor(Spire::Common::Color::GetBlack()), 0.1f);
	page->GetCanvas()->DrawRectangle(pen, new RectangleF(new PointF(x, y), new SizeF(width, height)));
	y = y + height - 50;
	width = 100;
	height = 50;
	//Initialize an instance of PdfSeparationColorSpace
	intrusive_ptr<PdfSeparationColorSpace> cs = new PdfSeparationColorSpace(L"MyColor", new PdfRGBColor(Spire::Common::Color::FromArgb(0, 100, 0, 0)));
	intrusive_ptr<PdfPen> pen1 = new PdfPen(new PdfRGBColor(Spire::Common::Color::GetRed()), 1.f);
	intrusive_ptr<PdfBrush> brush = new PdfSolidBrush(new PdfSeparationColor(cs, 0.1f));
	page->GetCanvas()->DrawRectangle(pen1, brush, new RectangleF(new PointF(x, y), new SizeF(width, height)));
	page->GetCanvas()->Restore(state);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

