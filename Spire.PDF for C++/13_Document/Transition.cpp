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
	wstring 	ImgFile = DATAPATH"Wikipedia_Science.png";
	intrusive_ptr<PdfImage> image = PdfImage::FromFile(ImgFile.c_str());
	page->GetCanvas()->DrawImage(image, new PointF(pageWidth - image->GetPhysicalDimension()->GetWidth(), y));
	float imageLeftSpace = pageWidth - image->GetPhysicalDimension()->GetWidth() - 2;
	float imageBottom = image->GetPhysicalDimension()->GetHeight() + y;

	intrusive_ptr<PdfTrueTypeFont> font3 = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
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

int main()
{
	wstring outputFile = OUTPUTPATH"Transition.pdf";
	//Create a pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->GetViewerPreferences()->SetPageMode(PdfPageMode::FullScreen);

	//Set the margin
	intrusive_ptr<PdfUnitConvertor> unitCvtr = new PdfUnitConvertor();
	intrusive_ptr<PdfMargins> margin = new PdfMargins();
	margin->SetTop(unitCvtr->ConvertUnits(2.54f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetBottom(margin->GetTop());
	margin->SetLeft(unitCvtr->ConvertUnits(3.17f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetRight(margin->GetLeft());

	//Create one section
	intrusive_ptr<PdfSection> section = doc->GetSections()->Add();
	section->GetPageSettings()->SetSize(PdfPageSize::A4());
	section->GetPageSettings()->SetMargins(margin);
	section->GetPageSettings()->SetTransition(new PdfPageTransition());
	section->GetPageSettings()->GetTransition()->SetDuration(2);
	section->GetPageSettings()->GetTransition()->SetStyle(PdfTransitionStyle::Fly);
	section->GetPageSettings()->GetTransition()->SetPageDuration(1);

	intrusive_ptr<PdfNewPage> page = section->GetPages()->Add();
	page->SetBackgroundColor(Color::GetRed());
	DrawPage(page);

	page = section->GetPages()->Add();
	page->SetBackgroundColor(Color::GetGreen());
	DrawPage(page);

	page = section->GetPages()->Add();
	page->SetBackgroundColor(Color::GetBlue());
	DrawPage(page);

	section = doc->GetSections()->Add();
	section->GetPageSettings()->SetSize(PdfPageSize::A4());
	section->GetPageSettings()->SetMargins(margin);
	section->GetPageSettings()->SetTransition(new PdfPageTransition());
	section->GetPageSettings()->GetTransition()->SetDuration(2);
	section->GetPageSettings()->GetTransition()->SetStyle(PdfTransitionStyle::Box);
	section->GetPageSettings()->GetTransition()->SetPageDuration(1);

	page = section->GetPages()->Add();
	page->SetBackgroundColor(Color::GetOrange());
	DrawPage(page);

	page = section->GetPages()->Add();
	page->SetBackgroundColor(Color::GetBrown());
	DrawPage(page);

	page = section->GetPages()->Add();
	page->SetBackgroundColor(Color::GetNavy());
	DrawPage(page);

	section = doc->GetSections()->Add();
	section->GetPageSettings()->SetSize(PdfPageSize::A4());
	section->GetPageSettings()->SetMargins(margin);
	section->GetPageSettings()->SetTransition(new PdfPageTransition());
	section->GetPageSettings()->GetTransition()->SetDuration(2);
	section->GetPageSettings()->GetTransition()->SetStyle(PdfTransitionStyle::Split);
	section->GetPageSettings()->GetTransition()->SetDimension(PdfTransitionDimension::Vertical);
	section->GetPageSettings()->GetTransition()->SetMotion(PdfTransitionMotion::Inward);
	section->GetPageSettings()->GetTransition()->SetPageDuration(1);

	page = section->GetPages()->Add();
	page->SetBackgroundColor(Color::GetOrange());
	DrawPage(page);

	page = section->GetPages()->Add();
	page->SetBackgroundColor(Color::GetBrown());
	DrawPage(page);

	page = section->GetPages()->Add();
	page->SetBackgroundColor(Color::GetNavy());
	DrawPage(page);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
