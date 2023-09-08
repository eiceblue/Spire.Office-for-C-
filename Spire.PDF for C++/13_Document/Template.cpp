#include "pch.h"
#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

static std::wstring string_to_wstring(const std::string& str)
{
	std::string strLocale = setlocale(LC_ALL, "");
	const char* chSrc = str.c_str();
	size_t nDestSize = mbstowcs(NULL, chSrc, 0) + 1;
	wchar_t* wchDest = new wchar_t[nDestSize];
	wmemset(wchDest, 0, nDestSize);
	mbstowcs(wchDest, chSrc, nDestSize);
	std::wstring wstrResult = wchDest;
	delete[]wchDest;
	setlocale(LC_ALL, strLocale.c_str());
	return wstrResult;
}

class StringBuilder
{
private:
	std::wstring privateString;

public:
	StringBuilder()
	{
	}

	StringBuilder* appendLine(const std::wstring& toAppend)
	{
		privateString += toAppend + L"\r\n";
		return this;
	}
	std::wstring toString()
	{
		return privateString;
	}
};

void DrawPage(intrusive_ptr<PdfPageBase> page)
{
	float pageWidth = page->GetCanvas()->GetClientSize()->GetWidth();
	float y = 0;

	//Title
	y = y + 5;
	intrusive_ptr<PdfBrush> brush2 = new PdfSolidBrush(new PdfRGBColor(Color::GetBlack()));

	intrusive_ptr<PdfTrueTypeFont> font2 = new PdfTrueTypeFont(L"Arial", 16.f, PdfFontStyle::Bold, true);
	intrusive_ptr<PdfStringFormat> format2 = new PdfStringFormat(PdfTextAlignment::Center);
	format2->SetCharacterSpacing(1.0f);
	std::wstring text = L"Summary of Science";
	page->GetCanvas()->DrawString(text.c_str(), font2, brush2, pageWidth / 2, y, format2);
	intrusive_ptr<SizeF> size = font2->MeasureString(text.c_str(), format2);
	y = y + size->GetHeight() + 6;

	//Icon
	wstring ImgFile = DATAPATH"Wikipedia_Science.png";
	intrusive_ptr<PdfImage> image = PdfImage::FromFile(ImgFile.c_str());
	page->GetCanvas()->DrawImage(image, new PointF(pageWidth - image->GetPhysicalDimension()->GetWidth(), y));
	float imageLeftSpace = pageWidth - image->GetPhysicalDimension()->GetWidth() - 2;
	float imageBottom = image->GetPhysicalDimension()->GetHeight() + y;

	
	intrusive_ptr<PdfTrueTypeFont> font3 = new PdfTrueTypeFont(L"Arial", 9.f, PdfFontStyle::Regular, true);
	intrusive_ptr<PdfStringFormat> format3 = new PdfStringFormat();
	format3->SetParagraphIndent(font3->GetSize() * 2);
	format3->SetMeasureTrailingSpaces(true);
	format3->SetLineSpacing(font3->GetSize() * 1.5f);
	std::wstring text1 = L"(All text and picture from ";
	std::wstring text2 = L"Wikipedia";
	std::wstring text3 = L", the free encyclopedia)";
	page->GetCanvas()->DrawString(text1.c_str(), font3, brush2, 0, y, format3);

	size = font3->MeasureString(text1.c_str(), format3);
	float x1 = size->GetWidth();
	format3->SetParagraphIndent(0);
	
	intrusive_ptr<PdfTrueTypeFont> font4 = new PdfTrueTypeFont(L"Arial", 9.f, PdfFontStyle::Underline, true);
	intrusive_ptr<PdfBrush> brush3 = PdfBrushes::GetBlue();
	page->GetCanvas()->DrawString(text2.c_str(), font4, brush3, x1, y, format3);
	size = font4->MeasureString(text2.c_str(), format3);
	x1 = x1 + size->GetWidth();

	page->GetCanvas()->DrawString(text3.c_str(), font3, brush2, x1, y, format3);
	y = y + size->GetHeight();

	//Content
	intrusive_ptr<PdfStringFormat> format4 = new PdfStringFormat();
	//text = System::IO::File::ReadAllText(DataPath"Demo/Summary_of_Science.txt)");

	wstring inputFile = DATAPATH"Summary_of_Science.txt";

	ifstream in_file;
	in_file.open(inputFile.c_str(), ios::in);
	if (!in_file.is_open())
	{
		cout << "Fail to read txt file" << endl;
		return;
	}
	StringBuilder* sb = new StringBuilder();
	string buf;
	while (getline(in_file, buf))
	{
		sb->appendLine(string_to_wstring(buf));
	}
	text = sb->toString();
	
	intrusive_ptr<PdfTrueTypeFont> font5 = new PdfTrueTypeFont(L"Arial", 10.f, PdfFontStyle::Regular, true);
	format4->SetLineSpacing(font5->GetSize() * 1.5f);
	intrusive_ptr<PdfStringLayouter> textLayouter = new PdfStringLayouter();
	float imageLeftBlockHeight = imageBottom - y;
	intrusive_ptr<PdfStringLayoutResult> result = textLayouter->Layout(text.c_str(), font5, format4, new SizeF(imageLeftSpace, imageLeftBlockHeight));
	if (result->GetActualSize()->GetHeight() < imageBottom - y)
	{
		imageLeftBlockHeight = imageLeftBlockHeight + result->GetLineHeight();
		result = textLayouter->Layout(text.c_str(), font5, format4, new SizeF(imageLeftSpace, imageLeftBlockHeight));
	}

	std::vector<intrusive_ptr<LineInfo>> line = result->GetLines();
	for (int i = 0; i < line.size(); i++)
	{
		page->GetCanvas()->DrawString(line[i]->GetText(), font5, brush2, 0, y, format4);
		y = y + result->GetLineHeight();
	}

	intrusive_ptr<PdfTextWidget> textWidget = new PdfTextWidget(result->GetRemainder(), font5, brush2);
	intrusive_ptr<PdfTextLayout> textLayout = new PdfTextLayout();
	textLayout->SetBreak(PdfLayoutBreakType::FitPage);
	textLayout->SetLayout(PdfLayoutType::Paginate);
	intrusive_ptr<RectangleF> bounds = new RectangleF(new PointF(0, y), page->GetCanvas()->GetClientSize());
	textWidget->SetStringFormat(format4);
	textWidget->Draw(Object::Convert<PdfNewPage>(page), bounds, textLayout);
}

void SetDocumentTemplate(intrusive_ptr<PdfDocument> doc, intrusive_ptr < SizeF> pageSize, intrusive_ptr<PdfMargins> margin)
{
	intrusive_ptr<PdfPageTemplateElement> leftSpace = new PdfPageTemplateElement(margin->GetLeft(), pageSize->GetHeight());
	doc->GetTemplate()->SetLeft(leftSpace);

	intrusive_ptr<PdfPageTemplateElement> topSpace = new PdfPageTemplateElement(pageSize->GetWidth(), margin->GetTop());
	topSpace->SetForeground(true);
	doc->GetTemplate()->SetTop(topSpace);

	
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 9.f, PdfFontStyle::Italic, true);
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat(PdfTextAlignment::Right);
	std::wstring label = L"Demo of Spire.Pdf";
	intrusive_ptr<SizeF> size = font->MeasureString(label.c_str(), format);
	float y = topSpace->GetHeight() - font->GetHeight() - 1;
	intrusive_ptr<PdfPen> pen = new PdfPen(new PdfRGBColor(Color::GetBlack()), 0.75f);
	topSpace->GetGraphics()->SetTransparency(0.5f);
	topSpace->GetGraphics()->DrawLine(pen, margin->GetLeft(), y, pageSize->GetWidth() - margin->GetRight(), y);
	y = y - 1 - size->GetHeight();
	topSpace->GetGraphics()->DrawString(label.c_str(), font, PdfBrushes::GetBlack(), pageSize->GetWidth() - margin->GetRight(), y, format);

	intrusive_ptr<PdfPageTemplateElement> rightSpace = new PdfPageTemplateElement(margin->GetRight(), pageSize->GetHeight());
	doc->GetTemplate()->SetRight(rightSpace);

	intrusive_ptr<PdfPageTemplateElement> bottomSpace = new PdfPageTemplateElement(pageSize->GetWidth(), margin->GetBottom());
	bottomSpace->SetForeground(true);
	doc->GetTemplate()->SetBottom(bottomSpace);

	//Draw footer label
	y = font->GetHeight() + 1;
	bottomSpace->GetGraphics()->SetTransparency(0.5f);
	bottomSpace->GetGraphics()->DrawLine(pen, margin->GetLeft(), y, pageSize->GetWidth() - margin->GetRight(), y);
	y = y + 1;
	intrusive_ptr<PdfPageNumberField> pageNumber = new PdfPageNumberField();
	intrusive_ptr<PdfPageCountField> pageCount = new PdfPageCountField();
	intrusive_ptr<PdfCompositeField> pageNumberLabel = new PdfCompositeField();
	pageNumberLabel->SetAutomaticFields({ pageNumber, pageCount });
	pageNumberLabel->SetBrush(PdfBrushes::GetBlack());
	pageNumberLabel->SetFont(font);
	pageNumberLabel->SetStringFormat(format);
	pageNumberLabel->SetText(L"page {0} of {1}");
	pageNumberLabel->Draw(bottomSpace->GetGraphics(), pageSize->GetWidth() - margin->GetRight(), y);
	wstring ImgFile1 = DATAPATH"Header.png";
	intrusive_ptr<PdfImage> headerImage = PdfImage::FromFile(ImgFile1.c_str());
	intrusive_ptr<PointF> pageLeftTop = new PointF(-margin->GetLeft(), -margin->GetTop());
	intrusive_ptr<PdfPageTemplateElement> header = new PdfPageTemplateElement(pageLeftTop, headerImage->GetPhysicalDimension());
	header->SetForeground(false);
	header->GetGraphics()->SetTransparency(0.5f);
	header->GetGraphics()->DrawImage(headerImage, 0.f, 0.f);
	doc->GetTemplate()->GetStamps()->Add(header);

	wstring ImgFile2 = DATAPATH"Footer.png";
	intrusive_ptr<PdfImage> footerImage = PdfImage::FromFile(ImgFile2.c_str());
	y = pageSize->GetHeight() - footerImage->GetPhysicalDimension()->GetHeight();
	intrusive_ptr<PointF> footerLocation = new PointF(-margin->GetLeft(), y);
	intrusive_ptr<PdfPageTemplateElement> footer = new PdfPageTemplateElement(footerLocation, footerImage->GetPhysicalDimension());
	footer->SetForeground(false);
	footer->GetGraphics()->SetTransparency(0.5f);;
	footer->GetGraphics()->DrawImage(footerImage, 0.f, 0.f);
	doc->GetTemplate()->GetStamps()->Add(footer);
}
void SetSectionTemplate(intrusive_ptr<PdfSection> section, intrusive_ptr < SizeF> pageSize, intrusive_ptr<PdfMargins> margin, const std::wstring& label)
{
	intrusive_ptr<PdfPageTemplateElement> leftSpace = new PdfPageTemplateElement(margin->GetLeft(), pageSize->GetHeight());
	leftSpace->SetForeground(true);
	section->GetTemplate()->SetOddLeft(leftSpace);

	
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 9.f, PdfFontStyle::Italic, true);
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat(PdfTextAlignment::Center, PdfVerticalAlignment::Middle);
	float y = (pageSize->GetHeight() - margin->GetTop() - margin->GetBottom()) * (1 - 0.618f);
	intrusive_ptr<RectangleF> bounds = new RectangleF(10, y, margin->GetLeft() - 20, font->GetHeight() + 6);
	leftSpace->GetGraphics()->DrawRectangle(PdfBrushes::GetOrangeRed(), bounds);
	leftSpace->GetGraphics()->DrawString(label.c_str(), font, PdfBrushes::GetWhite(), bounds, format);

	intrusive_ptr<PdfPageTemplateElement> rightSpace = new PdfPageTemplateElement(margin->GetRight(), pageSize->GetHeight());
	rightSpace->SetForeground(true);
	section->GetTemplate()->SetEvenRight(rightSpace);
	bounds = new RectangleF(10, y, margin->GetRight() - 20, font->GetHeight() + 6);
	rightSpace->GetGraphics()->DrawRectangle(PdfBrushes::GetSaddleBrown(), bounds);
	rightSpace->GetGraphics()->DrawString(label.c_str(), font, PdfBrushes::GetWhite(), bounds, format);
}


int main()
{
	wstring outputFile = OUTPUTPATH"Template.pdf";

	//Create a pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	doc->GetViewerPreferences()->SetPageLayout(PdfPageLayout::TwoColumnLeft);

	//Set the margin
	intrusive_ptr<PdfUnitConvertor> unitCvtr = new PdfUnitConvertor();
	intrusive_ptr<PdfMargins> margin = new PdfMargins();
	margin->SetTop(unitCvtr->ConvertUnits(2.54f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetBottom(margin->GetTop());
	margin->SetLeft(unitCvtr->ConvertUnits(3.17f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetRight(margin->GetLeft());

	SetDocumentTemplate(doc, new SizeF(PdfPageSize::A4()), margin);

	//Create one section
	intrusive_ptr<PdfSection> section = doc->GetSections()->Add();
	section->GetPageSettings()->SetSize(new SizeF(PdfPageSize::A4()));
	section->GetPageSettings()->SetMargins(new PdfMargins(0));
	SetSectionTemplate(section, new SizeF(PdfPageSize::A4()), margin, L"Section 1");

	//Create one page
	intrusive_ptr<PdfNewPage> page = section->GetPages()->Add();

	//Draw page
	DrawPage(page);

	page = section->GetPages()->Add();
	DrawPage(page);

	page = section->GetPages()->Add();
	DrawPage(page);

	page = section->GetPages()->Add();
	DrawPage(page);

	page = section->GetPages()->Add();
	DrawPage(page);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
