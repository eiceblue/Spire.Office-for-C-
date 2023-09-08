#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring outputFile = output_path + L"WrapTextAroundImage.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	float pageWidth = page->GetCanvas()->GetClientSize()->GetWidth();
	float y = 0;
	y += 8;
	intrusive_ptr<PdfBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetBlack()));
	intrusive_ptr<PdfTrueTypeFont> font1 = new PdfTrueTypeFont(L"Arial", 20.f, PdfFontStyle::Bold, true);
	intrusive_ptr<PdfStringFormat> format1 = new PdfStringFormat(PdfTextAlignment::Center);
	format1->SetCharacterSpacing(1);
	wstring text = L"Spire.PDF for C++";
	page->GetCanvas()->DrawString(text.c_str(), font1, brush, pageWidth / 2, y, format1);
	intrusive_ptr<SizeF> size = font1->MeasureString(text.c_str(), format1);
	y = y + size->GetHeight() + 6;
	wstring imageIn = input_path +L"PdfImage.png";
	intrusive_ptr<PdfImage> image = PdfImage::FromFile(imageIn.c_str());
	page->GetCanvas()->DrawImage(image, new PointF(pageWidth - image->GetPhysicalDimension()->GetWidth(), y));
	float imageLeftSpace = pageWidth - image->GetPhysicalDimension()->GetWidth() - 2;
	float imageBottom = image->GetPhysicalDimension()->GetHeight() + y;
	intrusive_ptr<PdfStringFormat> format2 = new PdfStringFormat();

	ifstream in_file;
	
	wstring textIn = input_path + L"text.txt";

	ifstream in(textIn.c_str(), ios::in);
	istreambuf_iterator<char> begin(in), end;
	wstring filetext(begin, end);
	in.close();

	intrusive_ptr<PdfTrueTypeFont> font2 = new PdfTrueTypeFont(L"Arial", 16.f, PdfFontStyle::Regular, true);
	format2->SetLineSpacing(font2->GetSize() * 1.5);
	intrusive_ptr<PdfStringLayouter> textLayouter = new PdfStringLayouter();
	float imageLeftBlockHeight = imageBottom - y;
	intrusive_ptr<PdfStringLayoutResult> result
		= textLayouter->Layout(filetext.c_str(), font2, format2, new SizeF(imageLeftSpace, imageLeftBlockHeight));
	if (result->GetActualSize()->GetHeight() < imageLeftBlockHeight)
	{
		imageLeftBlockHeight = imageLeftBlockHeight + result->GetLineHeight();
		result = textLayouter->Layout(filetext.c_str(), font2, format2, new SizeF(imageLeftSpace, imageLeftBlockHeight));
	}
	for(intrusive_ptr<LineInfo> line : result->GetLines())
	{
		page->GetCanvas()->DrawString(line->GetText(), font2, brush, 0, y, format2);
		y = y + result->GetLineHeight();
	}
	intrusive_ptr<PdfTextWidget> textWidget = new PdfTextWidget(result->GetRemainder(), font2, brush);
	intrusive_ptr<PdfTextLayout> textLayout = new PdfTextLayout();
	textLayout->SetBreak(PdfLayoutBreakType::FitPage);
	textLayout->SetLayout(PdfLayoutType::Paginate);
	intrusive_ptr<RectangleF> bounds = new RectangleF(new PointF(0, y), page->GetCanvas()->GetClientSize());
	textWidget->SetStringFormat(format2);
	intrusive_ptr<PdfNewPage> newPage = Object::Convert<PdfNewPage>(page);
	textWidget->Draw(newPage, bounds, textLayout);
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

	