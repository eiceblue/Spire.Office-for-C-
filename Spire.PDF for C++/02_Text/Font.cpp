#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path+L"Font.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"Font.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	float l = page->GetCanvas()->GetClientSize()->GetWidth() / 2;
	intrusive_ptr<PointF> center = new PointF(l, l);
	float r = sqrt(2 * l * l);
	intrusive_ptr<PdfRadialGradientBrush> brush = new PdfRadialGradientBrush(center, 0.f, center, r,
		new PdfRGBColor(Color::GetBlue()), new PdfRGBColor(Color::GetRed()));
	//Helvetica
	float x1 = 40;
	float y = 200;
	wstring text1 = L"Font Family: Helvetica";
	intrusive_ptr<PdfFont> font1 = new PdfFont(PdfFontFamily::Courier, 14.f);
	intrusive_ptr<PdfFont> font2 = new PdfFont(PdfFontFamily::Helvetica, 14.f);
	float x2 = x1 + 10 + font1->MeasureString(text1.c_str())->GetWidth();
	page->GetCanvas()->DrawString(text1.c_str(), font1, brush, x1, y);
	page->GetCanvas()->DrawString(text1.c_str(), font2, brush, x2, y);

	//Courier
	x1 = 40;
	y += 16;
	wstring text2 = L"Font Family: Courier";
	//intrusive_ptr<PdfFont> font1 = new PdfFont(PdfFontFamily::Courier, 14.f);
	font2 = new PdfFont(PdfFontFamily::Courier, 14.f);
	x2 = x1 + 10 + font1->MeasureString(text2.c_str())->GetWidth();
	page->GetCanvas()->DrawString(text2.c_str(), font1, brush, x1, y);
	page->GetCanvas()->DrawString(text2.c_str(), font2, brush, x2, y);

	//TimesRoman
	x1 = 40;
	y += 16;
	wstring text3 = L"Font Family: TimesRoman";
	//intrusive_ptr<PdfFont> font1 = new PdfFont(PdfFontFamily::Courier, 14.f);
	font2 = new PdfFont(PdfFontFamily::TimesRoman, 14.f);
	x2 = x1 + 10 + font1->MeasureString(text3.c_str())->GetWidth();
	page->GetCanvas()->DrawString(text3.c_str(), font1, brush, x1, y);
	page->GetCanvas()->DrawString(text3.c_str(), font2, brush, x2, y);

	//Symbol
	x1 = 40;
	y += 16;
	wstring text4 = L"Font Family: Symbol";
	//intrusive_ptr<PdfFont> font1 = new PdfFont(PdfFontFamily::Courier, 14.f);
	font2 = new PdfFont(PdfFontFamily::Symbol, 14.f);
	x2 = x1 + 10 + font1->MeasureString(text4.c_str())->GetWidth();
	page->GetCanvas()->DrawString(text4.c_str(), font1, brush, x1, y);
	page->GetCanvas()->DrawString(text4.c_str(), font2, brush, x2, y);

	//ZapfDingbats
	x1 = 40;
	y += 16;
	wstring text5 = L"Font Family: ZapfDingbats";
	//intrusive_ptr<PdfFont> font1 = new PdfFont(PdfFontFamily::Courier, 14.f);
	font2 = new PdfFont(PdfFontFamily::Symbol, 14.f);
	x2 = x1 + 10 + font1->MeasureString(text5.c_str())->GetWidth();
	page->GetCanvas()->DrawString(text5.c_str(), font1, brush, x1, y);
	page->GetCanvas()->DrawString(text5.c_str(), font2, brush, x2, y);

	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 15.f, PdfFontStyle::Bold, true);
	page->GetCanvas()->DrawString(L"Font Family: Arial - Embedded", font, brush, 40, (y = y + 26.f));
	wstring arabicText
		= L"\u0627\u0644\u0630\u0647\u0627\u0628\u0021\u0020\u0628\u062F\u0648\u0631\u0647\u0020\u062D\u0648\u0644\u0647\u0627\u0021\u0020\u0627\u0644\u0630\u0647\u0627\u0628\u0021\u0020\u0627\u0644\u0630\u0647\u0627\u0628\u0021\u0020\u0627\u0644\u0630\u0647\u0627\u0628\u0021";
	intrusive_ptr<RectangleF> rectg = new RectangleF(new PointF(40, (y = y + 26.f)), page->GetCanvas()->GetClientSize());
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat(PdfTextAlignment::Right);
	format->SetRightToLeft(true);
	page->GetCanvas()->DrawString(arabicText.c_str(), font, brush, rectg, format);

	font = new PdfTrueTypeFont(L"Batang", 14.f, PdfFontStyle::Italic, true);
	page->GetCanvas()->DrawString(L"Font Family: Batang - Not Embedded", font, brush, 40, (y = y + 16.f));
	wstring fontFileName = input_path+L"PT_Serif-Caption-Web-Regular.ttf";
	font = new PdfTrueTypeFont(fontFileName.c_str(), 20.f);
	page->GetCanvas()->DrawString(L"PT Serif Caption Font", font, brush, 40, (y = y + 36.f));
	page->GetCanvas()->DrawString(L"PT Serif Caption Font",
		new PdfFont(PdfFontFamily::Helvetica, 8.f), brush, 40, (y = y + 40.f));
	intrusive_ptr<PdfCjkStandardFont> cjkFont = new PdfCjkStandardFont(PdfCjkFontFamily::MonotypeHeiMedium, 14.f);
	page->GetCanvas()->DrawString(L"How to say 'Font' in Chinese? \u5B57\u4F53", cjkFont, brush, 40, (y = y + 36.f));
	cjkFont = new PdfCjkStandardFont(PdfCjkFontFamily::HanyangSystemsGothicMedium, 14.f);
	page->GetCanvas()->DrawString(L"How to say 'Font' in Japanese? \u30D5\u30A9\u30F3\u30C8", cjkFont, brush, 40, (y = y + 36.f));
	cjkFont = new PdfCjkStandardFont(PdfCjkFontFamily::HanyangSystemsShinMyeongJoMedium, 14.f);
	page->GetCanvas()->DrawString(L"How to say 'Font' in Korean? \uAE00\uAF34", cjkFont, brush, 40, (y = y + 36.f));
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}


