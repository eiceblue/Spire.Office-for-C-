#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DrawingTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DrawFilledRectangles.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfGraphicsState> state = page->GetCanvas()->Save();
	int x = 200;
	int y = 300;
	int width = 200;
	int height = 120;
	intrusive_ptr<PdfPen> pen = new PdfPen(new PdfRGBColor(Spire::Common::Color::GetBlack()), 1.f);
	intrusive_ptr<PdfBrush> brush = new PdfSolidBrush(new PdfRGBColor(Spire::Common::Color::GetOrangeRed()));
	page->GetCanvas()->DrawRectangle(pen, brush, new RectangleF(new PointF(x, y), new SizeF(width, height)));
	page->GetCanvas()->Restore(state);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

