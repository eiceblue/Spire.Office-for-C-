#include "../pch.h"

using namespace Spire::Pdf;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile =input_path+L"SampleB_1.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path +L"AddBorderForText.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile((inputFile).c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	LPCWSTR_S text = L"Hello, World!";
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Times New Roman", 11, PdfFontStyle::Regular, true);
	intrusive_ptr<SizeF> size = font->MeasureString(text);
	intrusive_ptr<PdfRGBColor> color = new PdfRGBColor(Color::GetBlack());
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(color);
	int x = 60;
	int y = 300;
	page->GetCanvas()->DrawString(text, font, brush, x, y);
	page->GetCanvas()->DrawRectangle(new PdfPen(brush, 0.5f), new Spire::Common::RectangleF(
		x, y, (int)size->GetWidth(), (int)size->GetHeight()));
	doc->SaveToFile((outputFile).c_str(), FileFormat::PDF);
	doc->Close();
}

