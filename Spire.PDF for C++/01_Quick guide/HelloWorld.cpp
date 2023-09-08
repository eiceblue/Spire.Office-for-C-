#include "../pch.h"

using namespace Spire::Pdf;

int main() {
	std::wstring output_path = OUTPUTPATH;
	std::wstring outputFile = output_path + L"HelloWorld.pdf";
	

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();

	
	std::wstring s = L"Hello, World!";
	float x = 10;
	float y = 10;
	intrusive_ptr<PdfFontBase> font = new PdfFont(PdfFontFamily::Helvetica, 30.f);
	intrusive_ptr<PdfRGBColor> color = new PdfRGBColor(Spire::Common::Color::GetBlack());
	intrusive_ptr<PdfBrush> textBrush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawString(s.c_str(), font, textBrush, x, y);
	
	doc->SaveToFile((outputFile).c_str());
	doc->Dispose();
}


