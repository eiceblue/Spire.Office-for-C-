#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DrawContentWithSpotColor.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DrawContentWithSpotColor.pdf";


	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfSeparationColorSpace> cs = new PdfSeparationColorSpace(L"MySpotColor", new PdfRGBColor(Spire::Common::Color::GetDarkViolet()));
	intrusive_ptr<PdfSeparationColor> color = new PdfSeparationColor(cs, 1.f);
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(color);

	//Draw pie with spot color(DarkViolet)
	page->GetCanvas()->DrawString(L"Tint=1.0", new PdfFont(PdfFontFamily::Helvetica, 10.f), brush, new PointF(160, 160));
	page->GetCanvas()->DrawPie(brush, 148, 200, 60, 60, 360, 360);

	page->GetCanvas()->DrawString(L"Tint=0.7", new PdfFont(PdfFontFamily::Helvetica, 10.f), brush, new PointF(230, 160));
	color = new PdfSeparationColor(cs, 0.7f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 218, 200, 60, 60, 360, 360);

	page->GetCanvas()->DrawString(L"Tint=0.4", new PdfFont(PdfFontFamily::Helvetica, 10.f), brush, new PointF(300, 160));
	color = new PdfSeparationColor(cs, 0.4f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 288, 200, 60, 60, 360, 360);

	page->GetCanvas()->DrawString(L"Tint=0.1", new PdfFont(PdfFontFamily::Helvetica, 10.f), brush, new PointF(370, 160));
	color = new PdfSeparationColor(cs, 0.1f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 358, 200, 60, 60, 360, 360);

	//Draw pie with spot color(Purple)
	cs = new PdfSeparationColorSpace(L"MySpotColor", new PdfRGBColor(Spire::Common::Color::GetPurple()));
	color = new PdfSeparationColor(cs, 1.f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 148, 280, 60, 60, 360, 360);

	color = new PdfSeparationColor(cs, 0.7f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 218, 280, 60, 60, 360, 360);

	color = new PdfSeparationColor(cs, 0.4f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 288, 280, 60, 60, 360, 360);

	color = new PdfSeparationColor(cs, 0.1f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 358, 280, 60, 60, 360, 360);

	//Draw pie with spot color(DarkSlateBlue)
	cs = new PdfSeparationColorSpace(L"MySpotColor", new PdfRGBColor(Spire::Common::Color::GetDarkSlateBlue()));
	color = new PdfSeparationColor(cs, 1.f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 148, 360, 60, 60, 360, 360);

	color = new PdfSeparationColor(cs, 0.7f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 218, 360, 60, 60, 360, 360);

	color = new PdfSeparationColor(cs, 0.4f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 288, 360, 60, 60, 360, 360);

	color = new PdfSeparationColor(cs, 0.1f);
	brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawPie(brush, 358, 360, 60, 60, 360, 360);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}


