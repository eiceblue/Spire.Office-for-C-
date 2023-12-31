#include "pch.h"

using namespace Spire::Common;
using namespace Spire::Pdf;
using namespace std;

class StringBuilder
{
private:
	std::wstring privateString;

public:
	StringBuilder()
	{
	}

	StringBuilder* append(const std::wstring& toAppend)
	{
		privateString += toAppend;
		return this;
	}

	std::wstring toString()
	{
		return privateString;
	}
};

int main()
{
	wstring outputFile = OUTPUTPATH"Link.pdf";
	//Create a pdf document.
	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	//Set margins
	intrusive_ptr<PdfUnitConvertor> unitCvtr = new PdfUnitConvertor();
	intrusive_ptr<PdfMargins> margin = new PdfMargins();
	margin->SetTop(unitCvtr->ConvertUnits(2.54f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetBottom(margin->GetTop());
	margin->SetLeft(unitCvtr->ConvertUnits(3.17f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetRight(margin->GetLeft());

	//Create one page
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add(new SizeF(PdfPageSize::A4()), margin);

	float y = 100;
	float x = 10;

	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Lucida Sans Unicode", 14, PdfFontStyle::Regular, true);
	std::wstring label = L"Simple Text Link: ";
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
	format->SetMeasureTrailingSpaces(true);
	page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrange(), 0, y, format);
	x = font->MeasureString(label.c_str(), format)->GetWidth();

	intrusive_ptr<PdfTrueTypeFont> font1 = new PdfTrueTypeFont(L"Lucida Sans Unicode", 14, PdfFontStyle::Underline, true);
	std::wstring url1 = L"http://www.e-iceblue.com";
	page->GetCanvas()->DrawString(url1.c_str(), font1, PdfBrushes::GetCadetBlue(), x, y);
	y = y + font1->MeasureString(url1.c_str())->GetHeight() + 25;

	label = L"Web Link: ";
	page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrange(), 0, y, format);
	x = font->MeasureString(label.c_str(), format)->GetWidth();
	std::wstring text = L"E-iceblue home";
	intrusive_ptr<PdfTextWebLink> link2 = new PdfTextWebLink();
	link2->SetText(text.c_str());
	link2->SetUrl(url1.c_str());
	link2->SetFont(font1);
	link2->SetBrush(PdfBrushes::GetCadetBlue());
	link2->DrawTextWebLink(page->GetCanvas(), new PointF(x, y));
	y = y + font1->MeasureString(text.c_str())->GetHeight() + 30;

	label = L"URI Annonation: ";
	page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrange(), 0, y, format);
	x = font->MeasureString(label.c_str(), format)->GetWidth();
	text = L"Google";
	intrusive_ptr<PointF> location = new PointF(x, y);
	intrusive_ptr<SizeF> size = font1->MeasureString(text.c_str());
	intrusive_ptr<RectangleF> linkBounds = new RectangleF(location, size);
	intrusive_ptr<PdfUriAnnotation> link3 = new PdfUriAnnotation(linkBounds);
	link3->SetBorder(new PdfAnnotationBorder(0));
	link3->SetUri(L"http://www.google.com");
	intrusive_ptr<PdfNewPage> newPage = Object::Convert<PdfNewPage>(page);
	newPage->GetAnnotations()->Add(link3);
	page->GetCanvas()->DrawString(text.c_str(), font1, PdfBrushes::GetCadetBlue(), x, y);
	y = y + size->GetHeight() + 30;

	label = L"URI Annonation Action: ";
	page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrange(), 0, y, format);
	x = font->MeasureString(label.c_str(), format)->GetWidth();
	text = L"JavaScript Action (Click Me)";
	location = new PointF(x - 2, y - 2);
	size = font1->MeasureString(text.c_str());
	size = new SizeF(size->GetWidth() + 5, size->GetHeight() + 5);
	linkBounds = new RectangleF(location, size);
	intrusive_ptr<PdfUriAnnotation> link4 = new PdfUriAnnotation(linkBounds);
	link4->SetBorder(new PdfAnnotationBorder(0.75f));
	link4->SetColor(new PdfRGBColor(Color::GetCadetBlue()));
	//Script
	StringBuilder* sb = new StringBuilder();
	sb->append(L"app.alert({");
	sb->append(L"    cMsg: \"Hello.\",");
	sb->append(L"    nIcon: 3,");
	sb->append(L"    cTitle: \"JavaScript Action\"");
	sb->append(L"});");
	//std::wstring script = L"app.alert({" + L"    cMsg: \"Hello.\"," + L"    nIcon: 3," + L"    cTitle: \"JavaScript Action\"" + L"});";
	intrusive_ptr<PdfJavaScriptAction> action = new PdfJavaScriptAction(sb->toString().c_str());
	link4->SetAction(action);
	intrusive_ptr<PdfNewPage> newPage1 = Object::Convert<PdfNewPage>(page);
	newPage1->GetAnnotations()->Add(link4);
	page->GetCanvas()->DrawString(text.c_str(), font1, PdfBrushes::GetCadetBlue(), x, y);
	y = y + size->GetHeight() + 30;

	label = L"Need Help:  ";
	page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrange(), 0, y, format);
	x = font->MeasureString(label.c_str(), format)->GetWidth();
	text = L"Go to forum to ask questions";
	link2 = new PdfTextWebLink();
	link2->SetText(text.c_str());
	link2->SetUrl(L"https://www.e-iceblue.com/forum/components-f5.html");
	link2->SetFont(font1);
	link2->SetBrush(PdfBrushes::GetCadetBlue());
	link2->DrawTextWebLink(page->GetCanvas(), new PointF(x, y));
	y = y + font1->MeasureString(text.c_str())->GetHeight() + 30;

	label = L"Contct us:  ";
	page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrange(), 0, y, format);
	x = font->MeasureString(label.c_str(), format)->GetWidth();
	text = L"Send an email";
	link2 = new PdfTextWebLink();
	link2->SetText(text.c_str());
	link2->SetUrl(L"mailto:support@e-iceblue.com");
	link2->SetFont(font1);
	link2->SetBrush(PdfBrushes::GetCadetBlue());
	link2->DrawTextWebLink(page->GetCanvas(), new PointF(x, y));
	y = y + font1->MeasureString(text.c_str())->GetHeight() + 30;

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}