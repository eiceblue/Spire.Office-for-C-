#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;


void DrawSuperscript(intrusive_ptr<PdfPageBase> page, const std::wstring& text, intrusive_ptr<PdfTrueTypeFont> font, intrusive_ptr<PdfSolidBrush> brush)
{
	float x1 = 120;
	float y1 = 100;
	page->GetCanvas()->DrawString(text.c_str(), font, brush, new PointF(x1, y1));
	intrusive_ptr<SizeF> size1 = font->MeasureString(text.c_str());
	x1 += size1->GetWidth();
	intrusive_ptr<PdfStringFormat> format1 = new PdfStringFormat();
	format1->SetSubSuperScript(PdfSubSuperScript::SuperScript);
	page->GetCanvas()->DrawString(L"Superscript", font, brush, new PointF(x1, y1), format1);
}

void DrawSubscript(intrusive_ptr<PdfPageBase> page, const std::wstring& text, intrusive_ptr<PdfTrueTypeFont> font, intrusive_ptr<PdfSolidBrush> brush)
{
	float x2 = 120;
	float y2 = 150;
	page->GetCanvas()->DrawString(text.c_str(), font, brush, new PointF(x2, y2));
	intrusive_ptr<SizeF> size2 = font->MeasureString(text.c_str());
	x2 += size2->GetWidth();
	intrusive_ptr<PdfStringFormat> format2 = new PdfStringFormat();
	format2->SetSubSuperScript(PdfSubSuperScript::SubScript);
	page->GetCanvas()->DrawString(L"SubScript", font, brush, new PointF(x2, y2), format2);
}

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SuperScriptAndSubScript.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 20.f, PdfFontStyle::Regular, true);
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetBlack()));
	wstring text = L"Spire.PDF for CPP";

	//Draw Superscript
	DrawSuperscript(page, text, font, brush);

	//Draw Subscript
	DrawSubscript(page, text, font, brush);
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}