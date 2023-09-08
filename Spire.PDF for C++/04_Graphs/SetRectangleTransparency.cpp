#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DrawingTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetRectangleTransparency.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfGraphicsState> state = page->GetCanvas()->Save();
	int x = 200;
	int y = 300;
	int width = 200;
	int height = 100;
	intrusive_ptr<PdfPen> pen = new PdfPen(new PdfRGBColor(Spire::Common::Color::GetBlack()), 1.f);
	intrusive_ptr<PdfBrush> brush = new PdfSolidBrush(new PdfRGBColor(Spire::Common::Color::GetRed()));
	page->GetCanvas()->SetTransparency(0.5f, 0.5f);
	page->GetCanvas()->DrawRectangle(pen, brush, new RectangleF(new PointF(x, y), new SizeF(width, height)));
	x = x + width / 2;
	y = y - height / 2;
	page->GetCanvas()->SetTransparency(0.2f, 0.2f);
	page->GetCanvas()->DrawRectangle(pen, brush, new RectangleF(new PointF(x, y), new SizeF(width, height)));
	page->GetCanvas()->Restore(state);
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

