#include "../pch.h"

using namespace Spire::Pdf;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path +L"AddTooltipToText.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 15.f, PdfFontStyle::Regular, true);
	page->GetCanvas()->DrawString(L"Move the mouse cursor over the following text to display a tooltip",
		font, PdfBrushes::GetBlack(), new PointF(10, 20));

	LPCWSTR_S text1 = L"Your Office Development Master";
	intrusive_ptr<PdfTrueTypeFont> font1 = new PdfTrueTypeFont(L"Arial", 18, PdfFontStyle::Regular, true);
	intrusive_ptr<SizeF> sizef1 = font1->MeasureString(text1);
	intrusive_ptr<RectangleF> rec1 = new RectangleF(new PointF(100, 100), sizef1);
	page->GetCanvas()->DrawString(text1, font1, new PdfSolidBrush(new PdfRGBColor(Spire::Common::Color::GetBlue())), rec1);
	intrusive_ptr<PdfButtonField> field1 = new PdfButtonField(page, L"field1");
	field1->SetBounds(rec1);
	field1->SetToolTip(L"E-iceblue Co. Ltd.");
	field1->SetBorderWidth(0);
	intrusive_ptr<PdfRGBColor> color = new PdfRGBColor(Color::GetTransparent());
	field1->SetBackColor(color);
	field1->SetForeColor(color);
	field1->SetLayoutMode(PdfButtonLayoutMode::IconOnly);
	field1->GetIconLayout()->SetIsFitBounds(true);

	LPCWSTR_S text2 = L"Spire.PDF";
	intrusive_ptr<PdfFont> font2 = new PdfFont(PdfFontFamily::TimesRoman, 20);
	intrusive_ptr<SizeF> sizef2 = font2->MeasureString(text2);
	intrusive_ptr<RectangleF> rec2 = new RectangleF(new PointF(100, 160), sizef2);
	page->GetCanvas()->DrawString(text2, font2, PdfBrushes::GetDarkOrange(), rec2);
	intrusive_ptr<PdfButtonField> field2 = new PdfButtonField(page, L"field2");
	field2->SetBounds(rec2);
	field2->SetToolTip(L"A professional PDF library applied to creating,writing, editing, handling and reading PDF files without any external dependencies.");
	field2->SetBorderWidth(0);
	field2->SetBackColor(color);
	field2->SetForeColor(color);
	field2->SetLayoutMode(PdfButtonLayoutMode::IconOnly);
	field2->GetIconLayout()->SetIsFitBounds(true);

	doc->SetAllowCreateForm(true);
	doc->GetForm()->GetFields()->Add(field1);
	doc->GetForm()->GetFields()->Add(field2);
	doc->SaveToFile((outputFile).c_str());
	doc->Close();
}

