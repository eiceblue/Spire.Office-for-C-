#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

static wstring trimStart(wstring source, const wstring& trimChars = L" \t\n\r\v\f")
{
	return source.erase(0, source.find_first_not_of(trimChars));
}

static wstring trimEnd(wstring source, const wstring& trimChars = L" \t\n\r\v\f")
{
	return source.erase(source.find_last_not_of(trimChars) + 1);
}

static std::wstring trim(std::wstring source, const std::wstring& trimChars = L" \t\n\r\v\f")
{
	return trimStart(trimEnd(source, trimChars), trimChars);
}

template<typename T>
static std::wstring formatSimple(const std::wstring& input, T arg)
{
	std::wostringstream ss;
	std::size_t lastCloseBrace = std::wstring::npos;
	std::size_t openBrace = std::wstring::npos;
	while ((openBrace = input.find(L'{', openBrace + 1)) != std::wstring::npos)
	{
		if (openBrace + 1 < input.length())
		{
			if (input[openBrace + 1] == L'{')
			{
				openBrace++;
				continue;
			}

			std::size_t closeBrace = input.find(L'}', openBrace + 1);
			if (closeBrace != std::wstring::npos)
			{
				ss << input.substr(lastCloseBrace + 1, openBrace - lastCloseBrace - 1);
				lastCloseBrace = closeBrace;

				std::wstring index = trim(input.substr(openBrace + 1, closeBrace - openBrace - 1));
				if (index == L"0")
					ss << arg;
				else
					throw std::runtime_error("Only simple positional format specifiers are handled by the 'formatSimple' helper method.");
			}
		}
	}

	if (lastCloseBrace + 1 < input.length())
		ss << input.substr(lastCloseBrace + 1);

	return ss.str();
}

int main()
{
	wstring outputFile = OUTPUTPATH"AddTableOfContent.pdf";
	wstring inputFile = DATAPATH"AddTableOfContent.pdf";

	//open a pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	//get the page count
	int pageCount = doc->GetPages()->GetCount();

	//insert a blank page into the pdf document
	intrusive_ptr<PdfPageBase> tocPage = doc->GetPages()->Insert(0);

	//set title
	std::wstring title = L"Table Of Contents";

	intrusive_ptr<PdfTrueTypeFont> titleFont = new PdfTrueTypeFont(L"Arial", 20, PdfFontStyle::Bold, true);
	intrusive_ptr<PdfStringFormat> centerAlignment = new PdfStringFormat(PdfTextAlignment::Center, PdfVerticalAlignment::Middle);
	intrusive_ptr<PointF> location = new PointF(tocPage->GetCanvas()->GetClientSize()->GetWidth() / 2, titleFont->MeasureString(title.c_str())->GetHeight());
	tocPage->GetCanvas()->DrawString(title.c_str(), titleFont, PdfBrushes::GetCornflowerBlue(), location, centerAlignment);

	intrusive_ptr<PdfTrueTypeFont> titlesFont = new PdfTrueTypeFont(L"Arial", 14, PdfFontStyle::Regular, true);
	std::vector<std::wstring> titles(pageCount);
	for (int i = 0; i < titles.size(); i++)
	{
		titles[i] = formatSimple(L"This is page{0}", i + 1);
	}
	float y = titleFont->MeasureString(title.c_str())->GetHeight() + 10;
	float x = 0;

	for (int i = 1; i <= pageCount; i++)
	{
		std::wstring text = titles[i - 1];
		SizeF titleSize = titlesFont->MeasureString(text.c_str());

		intrusive_ptr<PdfPageBase> navigatedPage = doc->GetPages()->GetItem(i);

		std::wstring pageNumText = std::to_wstring(i + 1);
		SizeF pageNumTextSize = titlesFont->MeasureString(pageNumText.c_str());
		tocPage->GetCanvas()->DrawString(text.c_str(), titlesFont, PdfBrushes::GetCadetBlue(), 0, y);
		float dotLocation = titleSize.GetWidth() + 2 + x;
		float pageNumlocation = tocPage->GetCanvas()->GetClientSize()->GetWidth() - pageNumTextSize.GetWidth();
		for (float j = dotLocation; j < pageNumlocation; j++)
		{
			if (dotLocation >= pageNumlocation)
			{
				break;
			}
			tocPage->GetCanvas()->DrawString(L".", titlesFont, PdfBrushes::GetGray(), dotLocation, y);
			dotLocation += 3;
		}
		tocPage->GetCanvas()->DrawString(pageNumText.c_str(), titlesFont, PdfBrushes::GetCadetBlue(), pageNumlocation, y);

		//add TOC action
		location = new PointF(0, y);
		intrusive_ptr<RectangleF> titleBounds = new RectangleF(location, new SizeF(tocPage->GetCanvas()->GetClientSize()->GetWidth(), titleSize.GetHeight()));
		intrusive_ptr<PdfDestination> Dest = new PdfDestination(navigatedPage, new PointF(-doc->GetPageSettings()->GetMargins()->GetTop(), -doc->GetPageSettings()->GetMargins()->GetLeft()));
		intrusive_ptr<PdfGoToAction> tempVar3 = new PdfGoToAction(Dest);
		intrusive_ptr<PdfActionAnnotation> action = new PdfActionAnnotation(titleBounds, tempVar3);
		action->SetBorder(new PdfAnnotationBorder(0));
		intrusive_ptr<PdfNewPage> newPage = Object::Convert<PdfNewPage>(tocPage);
		newPage->GetAnnotations()->Add(action);
		y += titleSize.GetHeight() + 10;

	}

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
