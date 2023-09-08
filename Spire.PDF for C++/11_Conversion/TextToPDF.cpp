#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

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

int main()
{
	wstring outputFile = OUTPUTPATH"TextToPDF.pdf";
	wstring inputFile = DATAPATH"TextToPdf.txt";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	ifstream in_file;
	in_file.open(inputFile.c_str());
	if (!in_file.is_open()) {
		cout << "It failed to read file" << endl;
		return -1;
	}
	StringBuilder* sb = new StringBuilder();
	string buf;
	while (getline(in_file, buf)) {
		sb->appendLine(string_to_wstring(buf));
	}
	wstring text = sb->toString();
	intrusive_ptr<PdfSection> section = doc->GetSections()->Add();
	intrusive_ptr<PdfPageBase> page = section->GetPages()->Add();
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 11);
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
	format->SetLineSpacing(20);
	intrusive_ptr<PdfBrush> brush = PdfBrushes::GetBlack();
	intrusive_ptr<PdfTextLayout> textLayout = new PdfTextLayout();
	textLayout->SetBreak(PdfLayoutBreakType::FitPage);
	textLayout->SetLayout(PdfLayoutType::Paginate);
	intrusive_ptr<RectangleF> bounds = new RectangleF(new PointF(10, 20), page->GetCanvas()->GetClientSize());
	intrusive_ptr<PdfTextWidget> textWidget = new PdfTextWidget(text.c_str(), font, brush);
	textWidget->SetStringFormat(format);
	intrusive_ptr<PdfNewPage> newPage = Object::Convert<PdfNewPage>(page);
	textWidget->Draw(newPage, bounds, textLayout);
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
