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
	LPCWSTR_S text = L"Summary of Science";
	page->GetCanvas()->DrawString(text, font2, brush2, pageWidth / 2, y, format2);
	intrusive_ptr<SizeF> size = font2->MeasureString(text, format2);
	y = y + size->GetHeight() + 6;

	//Icon

	wstring inputFile = DATAPATH"Wikipedia_Science.png";
	intrusive_ptr<PdfImage> image = PdfImage::FromFile(inputFile.c_str());
	page->GetCanvas()->DrawImage(image, new PointF(pageWidth - image->GetPhysicalDimension()->GetWidth(), y));
	float imageLeftSpace = pageWidth - image->GetPhysicalDimension()->GetWidth() - 2;
	float imageBottom = image->GetPhysicalDimension()->GetHeight() + y;

	//Reference content           
	intrusive_ptr<PdfTrueTypeFont> font3 = new PdfTrueTypeFont(L"Arial", 9.f, PdfFontStyle::Regular, true);

	intrusive_ptr<PdfStringFormat> format3 = new PdfStringFormat();
	format3->SetParagraphIndent(font3->GetSize() * 2);
	format3->SetMeasureTrailingSpaces(true);
	format3->SetLineSpacing(font3->GetSize() * 1.5f);
	LPCWSTR_S text1 = L"(All text and picture from ";
	LPCWSTR_S text2 = L"Wikipedia";
	LPCWSTR_S text3 = L", the free encyclopedia)";
	page->GetCanvas()->DrawString(text1, font3, brush2, 0, y, format3);

	size = font3->MeasureString(text1, format3);
	float x1 = size->GetWidth();
	format3->SetParagraphIndent(0);

	intrusive_ptr<PdfTrueTypeFont> font4 = new PdfTrueTypeFont(L"Arial", 9.0f, PdfFontStyle::Underline, true);

	intrusive_ptr<PdfBrush> brush3 = new PdfSolidBrush(new PdfRGBColor(Color::GetBlue()));
	page->GetCanvas()->DrawString(text2, font4, brush3, x1, y, format3);
	size = font4->MeasureString(text2, format3);
	x1 = x1 + size->GetWidth();

	page->GetCanvas()->DrawString(text3, font3, brush2, x1, y, format3);
	y = y + size->GetHeight();

	//Content
	intrusive_ptr<PdfStringFormat> format4 = new PdfStringFormat();

	wstring inputFile1 = DATAPATH"Summary_of_Science.txt";
	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	ifstream in_file;
	in_file.open(inputFile1.c_str(), ios::in);
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
	wstring text4 = sb->toString();

	LPCWSTR_S fontname = L"Arial";
	intrusive_ptr<PdfTrueTypeFont> font5 = new PdfTrueTypeFont(L"Arial", 10.0f, PdfFontStyle::Regular, true);

	format4->SetLineSpacing(font5->GetSize() * 1.5f);
	intrusive_ptr<PdfStringLayouter> textLayouter = new PdfStringLayouter();
	float imageLeftBlockHeight = imageBottom - y;
	intrusive_ptr<PdfStringLayoutResult> result = textLayouter->Layout(text4.c_str(), font5, format4, new SizeF(imageLeftSpace, imageLeftBlockHeight));
	if (result->GetActualSize()->GetHeight() < imageBottom - y)
	{
		imageLeftBlockHeight = imageLeftBlockHeight + result->GetLineHeight();
		result = textLayouter->Layout(text4.c_str(), font5, format4, new SizeF(imageLeftSpace, imageLeftBlockHeight));
	}
	if (page->GetImagesInfo().size() != 0)
	{
		vector< intrusive_ptr<LineInfo>> line = result->GetLines();
		for (int j = 0; j < line.size(); j++)
		{
			page->GetCanvas()->DrawString(line[j]->GetText(), font5, brush2, 0, y, format4);
			y = y + result->GetLineHeight();
		}
	}

	intrusive_ptr<PdfTextWidget> textWidget = new PdfTextWidget(result->GetRemainder(), font5, brush2);
	intrusive_ptr<PdfTextLayout> textLayout = new PdfTextLayout();
	textLayout->SetBreak(PdfLayoutBreakType::FitPage);
	textLayout->SetLayout(PdfLayoutType::Paginate);
	intrusive_ptr<RectangleF> bounds = new RectangleF(new PointF(0, y), page->GetCanvas()->GetClientSize());
	textWidget->SetStringFormat(format4);
	intrusive_ptr<PdfNewPage> newPage = Object::Convert<PdfNewPage>(page);
	textWidget->Draw(newPage, bounds, textLayout);
}

void CreatePDFA1WithSpirePDF_X1A2001()
{
	wstring outputFile = OUTPUTPATH"CreatePDFA1WithSpirePDF_X1A2001.pdf";
	wstring inputFile = OUTPUTPATH"Temp/CreatePDFA1WithSpirePDF_Temp.pdf";

	intrusive_ptr<Stream> stream = new Stream(inputFile.c_str());
	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(stream);
	converter->ToPdfX1A2001(outputFile.c_str());
	stream->Dispose();
	converter->Dispose();
}

void CreatePDFA1WithSpirePDF_A3B()
{
	wstring outputFile = OUTPUTPATH"CreatePDFA1WithSpirePDF_A3B.pdf";
	wstring inputFile = OUTPUTPATH"Temp/CreatePDFA1WithSpirePDF_Temp.pdf";

	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA3B(outputFile.c_str());
}

void CreatePDFA1WithSpirePDF_A3A()
{
	wstring outputFile = OUTPUTPATH"CreatePDFA1WithSpirePDF_A3A.pdf";
	wstring inputFile = OUTPUTPATH"Temp/CreatePDFA1WithSpirePDF_Temp.pdf";

	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA3A(outputFile.c_str());
}

void CreatePDFA1WithSpirePDF_A2A()
{
	wstring outputFile = OUTPUTPATH"CreatePDFA1WithSpirePDF_A2A.pdf";
	wstring inputFile = OUTPUTPATH"Temp/CreatePDFA1WithSpirePDF_Temp.pdf";

	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA2A(outputFile.c_str());
}

void CreatePDFA1WithSpirePDF_A2B()
{
	wstring outputFile = OUTPUTPATH"CreatePDFA1WithSpirePDF_A2B.pdf";
	wstring inputFile = OUTPUTPATH"Temp/CreatePDFA1WithSpirePDF_Temp.pdf";
	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA2B(outputFile.c_str());
}

void CreatePDFA1WithSpirePDF_A1A()
{
	wstring outputFile = OUTPUTPATH"CreatePDFA1WithSpirePDF_A1A.pdf";
	wstring inputFile = OUTPUTPATH"Temp/CreatePDFA1WithSpirePDF_Temp.pdf";

	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA1A(outputFile.c_str());
}

void CreatePDFA1WithSpirePDF_A1B()
{
	wstring outputFile = OUTPUTPATH"CreatePDFA1WithSpirePDF_A1B.pdf";
	wstring outputFile_Temp = OUTPUTPATH"Temp/CreatePDFA1WithSpirePDF_Temp.pdf";

	intrusive_ptr<PdfNewDocument> doc = new PdfNewDocument();
	intrusive_ptr<PdfMargins> tempVar = new PdfMargins(40);
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add(new SizeF(PdfPageSize::A4()), tempVar);

	// Draw content
	DrawPage(page);
	//Save the document
	intrusive_ptr<PdfDocumentBase> docBase = Object::Convert<PdfDocumentBase>(doc);
	docBase->Save(outputFile_Temp.c_str());
	doc->Close(true);
	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(outputFile_Temp.c_str());
	converter->ToPdfA1B(outputFile.c_str());
}

int main()
{
	CreatePDFA1WithSpirePDF_A1B();
	CreatePDFA1WithSpirePDF_A1A();
	CreatePDFA1WithSpirePDF_A2B();
	CreatePDFA1WithSpirePDF_A2A();
	CreatePDFA1WithSpirePDF_A3A();
	CreatePDFA1WithSpirePDF_A3B();
	CreatePDFA1WithSpirePDF_X1A2001();
}