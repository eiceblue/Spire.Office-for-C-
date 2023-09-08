#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring output_paths = OUTPUTPATH;
	wstring outputFile = output_paths + L"TextLayout.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	float pageWidth = page->GetCanvas()->GetClientSize()->GetWidth();
	float y = 0;

	//page header
	intrusive_ptr<PdfPen> pen1 = new PdfPen(new PdfRGBColor(Color::GetLightGray()), 1.f);
	intrusive_ptr<PdfBrush> brush1 = new PdfSolidBrush(new PdfRGBColor(Color::GetLightGray()));
	intrusive_ptr<PdfTrueTypeFont> font1 = new PdfTrueTypeFont(L"Arial", 8.f, PdfFontStyle::Italic, true);
	intrusive_ptr<PdfStringFormat> format1 = new PdfStringFormat(PdfTextAlignment::Right);
	wstring text = L"Demo of Spire.Pdf";
	page->GetCanvas()->DrawString(text.c_str(), font1, brush1, pageWidth - 2, y, format1);
	intrusive_ptr<SizeF> size = font1->MeasureString(text.c_str(), format1);
	y += size->GetHeight() + 1;
	page->GetCanvas()->DrawLine(pen1, 0, y, pageWidth, y);

	//title
	y += 25;
	intrusive_ptr<PdfBrush> brush2 = new PdfSolidBrush(new PdfRGBColor(Color::GetBlack()));
	intrusive_ptr<PdfTrueTypeFont> font2 = new PdfTrueTypeFont(L"Arial", 18.f, PdfFontStyle::Bold, true);
	intrusive_ptr<PdfStringFormat> format2 = new PdfStringFormat(PdfTextAlignment::Center);
	format2->SetCharacterSpacing(1.f);
	text = L"Summary of Science";
	page->GetCanvas()->DrawString(text.c_str(), font2, brush2, pageWidth / 2, y, format2);
	size = font2->MeasureString(text.c_str(), format2);
	y = y + size->GetHeight() + 16;

	//icon
	wstring inputFile = input_path +L"Wikipedia_Science.png";
	intrusive_ptr<PdfImage> image = PdfImage::FromFile(inputFile.c_str());
	page->GetCanvas()->DrawImage(image, new PointF(pageWidth - image->GetPhysicalDimension()->GetWidth(), y));
	float imageLeftSpace = pageWidth - image->GetPhysicalDimension()->GetWidth() - 2;
	float imageBottom = image->GetPhysicalDimension()->GetHeight() + y;
	intrusive_ptr<PdfTrueTypeFont> font3 = new PdfTrueTypeFont(L"Arial", 18.f, PdfFontStyle::Regular, true);
	intrusive_ptr<PdfStringFormat> format3 = new PdfStringFormat();
	format3->SetParagraphIndent(font3->GetSize() * 2);
	format3->SetMeasureTrailingSpaces(true);
	format3->SetLineSpacing(font3->GetSize() * 1.5);
	wstring text1 = L"(All text and picture from ";
	wstring text2 = L"Wikipedia";
	wstring text3 = L", the free encyclopedia)";
	page->GetCanvas()->DrawString(text1.c_str(), font3, brush2, 0, y, format3);
	size = font3->MeasureString(text1.c_str(), format3);
	float x1 = size->GetWidth();
	format3->SetParagraphIndent(0);
	intrusive_ptr<PdfTrueTypeFont> font4 = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Underline, true);
	intrusive_ptr<PdfBrush> brush3 = PdfBrushes::GetBlue();
	page->GetCanvas()->DrawString(text2.c_str(), font4, brush3, x1, y, format3);
	size = font3->MeasureString(text2.c_str(), format3);
	x1 += size->GetWidth();
	page->GetCanvas()->DrawString(text3.c_str(), font3, brush2, x1, y, format3);
	y += size->GetHeight();

	//Content
	PdfStringFormat* format4 = new PdfStringFormat();
	ifstream in_file;
	wstring textIn = input_path + L"Summary_of_Science.txt";

	ifstream in(textIn.c_str(), ios::in);
	istreambuf_iterator<char> begin(in), end;
	wstring filetext(begin, end);
	in.close();

	intrusive_ptr<PdfTrueTypeFont> font5 = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
	format4->SetLineSpacing(font5->GetSize() * 1.5);
	intrusive_ptr<PdfStringLayouter> textLayouter = new PdfStringLayouter();
	float imageLeftBlockHeight = imageBottom - y;
	intrusive_ptr<PdfStringLayoutResult> result
		= textLayouter->Layout(filetext.c_str(), font5, format4, new SizeF(imageLeftSpace, imageLeftBlockHeight));
	if (result->GetActualSize()->GetHeight() < imageBottom - y) {
		imageLeftBlockHeight = imageLeftBlockHeight + result->GetLineHeight();
		result = textLayouter->Layout(filetext.c_str(), font5, format4, new SizeF(imageLeftSpace, imageLeftBlockHeight));
	}
	for(intrusive_ptr<LineInfo> line : result->GetLines())
	{
		page->GetCanvas()->DrawString(line->GetText(), font5, brush2, 0, y, format4);
		y = y + result->GetLineHeight() + 2;
	}
	intrusive_ptr<PdfTextWidget> textWidget = new PdfTextWidget(result->GetRemainder(), font5, brush2);
	intrusive_ptr<PdfTextLayout> textLayout = new PdfTextLayout();
	textLayout->SetBreak(PdfLayoutBreakType::FitPage);
	textLayout->SetLayout(PdfLayoutType::Paginate);
	intrusive_ptr<RectangleF> bounds = new RectangleF(new PointF(0, y), page->GetCanvas()->GetClientSize());
	textWidget->SetStringFormat(format4);
	intrusive_ptr<PdfNewPage> newPage = Object::Convert<PdfNewPage>(page);
	textWidget->Draw(newPage, bounds, textLayout);
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
