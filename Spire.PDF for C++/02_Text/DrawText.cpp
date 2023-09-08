#include "../pch.h"

using namespace Spire::Pdf;
using namespace Spire::Common;

void AlignText(intrusive_ptr<PdfPageBase> page)
{
	//Draw the text - alignment
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 20.f);
	intrusive_ptr<PdfRGBColor> color = new PdfRGBColor(Color::GetBlue());
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(color);
	intrusive_ptr<PdfStringFormat> leftAlignment = new PdfStringFormat(PdfTextAlignment::Left, PdfVerticalAlignment::Middle);
	page->GetCanvas()->DrawString(L"Left!", font, brush, 0, 20, leftAlignment);
	page->GetCanvas()->DrawString(L"Left!", font, brush, 0, 50, leftAlignment);

	intrusive_ptr<PdfStringFormat> rightAlignment = new PdfStringFormat(PdfTextAlignment::Right, PdfVerticalAlignment::Middle);
	page->GetCanvas()->DrawString(L"Right!", font, brush, page->GetCanvas()->
		GetClientSize()->GetWidth(), 20, rightAlignment);
	page->GetCanvas()->DrawString(L"Right!", font, brush, page->GetCanvas()->
		GetClientSize()->GetWidth(), 50, rightAlignment);

	intrusive_ptr<PdfStringFormat> centerAlignment = new PdfStringFormat(PdfTextAlignment::Center, PdfVerticalAlignment::Middle);
	page->GetCanvas()->DrawString(L"Go! Turn Around! Go! Go! Go!", font, brush, page->GetCanvas()->
		GetClientSize()->GetWidth() / 2, 40, centerAlignment);
}

void AlignTextInRectangle(intrusive_ptr<PdfPageBase> page)
{
	//Draw the text - align in rectangle
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 10.f);
	intrusive_ptr<PdfRGBColor> color = new PdfRGBColor(Color::GetBlue());
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(color);
	intrusive_ptr<RectangleF> rectg1 = new RectangleF(0, 70, page->GetCanvas()->GetClientSize()->GetWidth() / 2, 100);
	intrusive_ptr<RectangleF> rectg2 = new RectangleF(page->GetCanvas()->GetClientSize()->GetWidth() / 2, 70,
		page->GetCanvas()->GetClientSize()->GetWidth() / 2, 100);
	page->GetCanvas()->DrawRectangle(new PdfSolidBrush(new PdfRGBColor(Color::GetLightBlue())), rectg1);
	page->GetCanvas()->DrawRectangle(new PdfSolidBrush(new PdfRGBColor(Color::GetLightSkyBlue())), rectg2);

	intrusive_ptr<PdfStringFormat> leftAlignment1 = new PdfStringFormat(PdfTextAlignment::Left, PdfVerticalAlignment::Top);
	page->GetCanvas()->DrawString(L"Left! Left!", font, brush, rectg1, leftAlignment1);
	page->GetCanvas()->DrawString(L"Left! Left!", font, brush, rectg2, leftAlignment1);

	intrusive_ptr<PdfStringFormat> rightAlignment1 = new PdfStringFormat(PdfTextAlignment::Right, PdfVerticalAlignment::Middle);
	page->GetCanvas()->DrawString(L"Right! Right!", font, brush, rectg1, rightAlignment1);
	page->GetCanvas()->DrawString(L"Right! Right!", font, brush, rectg2, rightAlignment1);

	intrusive_ptr<PdfStringFormat> centerAlignment1 = new PdfStringFormat(PdfTextAlignment::Center, PdfVerticalAlignment::Bottom);
	page->GetCanvas()->DrawString(L"Go! Turn Around! Go! Go! Go!", font, brush, rectg1, centerAlignment1);
	page->GetCanvas()->DrawString(L"Go! Turn Around! Go! Go! Go!", font, brush, rectg2, centerAlignment1);
}

void RotateText(intrusive_ptr<PdfPageBase> page)
{
	//Save graphics state
	intrusive_ptr<PdfGraphicsState> state = page->GetCanvas()->Save();
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 10.f);
	intrusive_ptr<PdfRGBColor> color = new PdfRGBColor(Color::GetBlue());
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(color);
	intrusive_ptr<PdfStringFormat> centerAlignment2 = new PdfStringFormat(PdfTextAlignment::Left, PdfVerticalAlignment::Middle);
	float x = page->GetCanvas()->GetClientSize()->GetWidth() / 2;
	float y = 380;
	page->GetCanvas()->TranslateTransform(x, y);
	for (int i = 0; i < 12; i++) {
		page->GetCanvas()->RotateTransform(30);
		page->GetCanvas()->DrawString(L"Go! Turn Around! Go! Go! Go!", font, brush, 20, 0, centerAlignment2);
	}
	page->GetCanvas()->Restore(state); 
}

void TransformText(intrusive_ptr<PdfPageBase> page)
{
	//Save graphics state
	intrusive_ptr<PdfGraphicsState> state = page->GetCanvas()->Save();
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 18.f);
	page->GetCanvas()->TranslateTransform(20, 200);
	page->GetCanvas()->ScaleTransform(1.f, 0.6f);
	page->GetCanvas()->SkewTransform(-10, 0);
	page->GetCanvas()->DrawString(L"Go! Turn Around! Go! Go! Go!", font, new PdfSolidBrush(new PdfRGBColor(Color::GetDeepSkyBlue())), 0.f, 0.f);
	page->GetCanvas()->SkewTransform(10, 0);
	page->GetCanvas()->DrawString(L"Go! Turn Around! Go! Go! Go!", font, new PdfSolidBrush(new PdfRGBColor(Color::GetCadetBlue())), 0.f, 0.f);
	page->GetCanvas()->ScaleTransform(1.f, -1.f);
	page->GetCanvas()->DrawString(L"Go! Turn Around! Go! Go! Go!", font, new PdfSolidBrush(new PdfRGBColor(Color::GetCadetBlue())), 0.f, -36.f);
	page->GetCanvas()->Restore(state);
}

void Draw_Text(intrusive_ptr<PdfPageBase> page)
{
	intrusive_ptr<PdfGraphicsState> state = page->GetCanvas()->Save();
	LPCWSTR_S text = L"Go! Turn Around! Go! Go! Go!";
	intrusive_ptr<PdfPen> pen = PdfPens::GetDeepSkyBlue();
	intrusive_ptr<PdfRGBColor> color1 = new PdfRGBColor(Color::GetWhite());
	intrusive_ptr<PdfBrush> brush = new PdfSolidBrush(color1);
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 18, PdfFontStyle::Italic);
	intrusive_ptr<SizeF> sizeF = font->MeasureString(text, format);
	intrusive_ptr<RectangleF> rctg = new RectangleF(page->GetCanvas()->GetClientSize()->GetWidth() / 2 + 10, 180,
		sizeF->GetWidth() / 3 * 2, sizeF->GetHeight() * 2);
	page->GetCanvas()->DrawString(text, font, pen, brush, rctg, format);
	page->GetCanvas()->Restore(state);
}

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DrawText.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();

	Draw_Text(page);
	AlignText(page);
	AlignTextInRectangle(page);
	TransformText(page);
	RotateText(page);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
