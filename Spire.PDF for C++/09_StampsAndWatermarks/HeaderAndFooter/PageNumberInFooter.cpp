#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

static wstring trimStart(wstring source, const wstring& trimChars = L" \t\n\r\v\f")
{
	return source.erase(0, source.find_first_not_of(trimChars));
}

static wstring trimEnd(wstring source, const wstring& trimChars = L" \t\n\r\v\f")
{
	return source.erase(source.find_last_not_of(trimChars) + 1);
}

static wstring trim(wstring source, const wstring& trimChars = L" \t\n\r\v\f")
{
	return trimStart(trimEnd(source, trimChars), trimChars);
}

template<typename T1, typename T2>
static wstring formatSimple(const wstring& input, T1 arg1, T2 arg2)
{
	wostringstream ss;
	size_t lastCloseBrace = wstring::npos;
	size_t openBrace = wstring::npos;
	while ((openBrace = input.find(L'{', openBrace + 1)) != wstring::npos)
	{
		if (openBrace + 1 < input.length())
		{
			if (input[openBrace + 1] == L'{')
			{
				openBrace++;
				continue;
			}

			size_t closeBrace = input.find(L'}', openBrace + 1);
			if (closeBrace != wstring::npos)
			{
				ss << input.substr(lastCloseBrace + 1, openBrace - lastCloseBrace - 1);
				lastCloseBrace = closeBrace;

				wstring index = trim(input.substr(openBrace + 1, closeBrace - openBrace - 1));
				if (index == L"0")
					ss << arg1;
				else if (index == L"1")
					ss << arg2;
				else
					throw runtime_error("Only simple positional format specifiers are handled by the 'formatSimple' helper method.");
			}
		}
	}

	if (lastCloseBrace + 1 < input.length())
		ss << input.substr(lastCloseBrace + 1);

	return ss.str();
}

void DrawPageNumber(intrusive_ptr<PdfDocument> doc, intrusive_ptr<PdfMargins> margin, int startNumber, int pageCount)
{

	for (int i = 0; i < doc->GetPages()->GetCount(); i++)
	{
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(i);
		page->GetCanvas()->SetTransparency(0.5f);
		intrusive_ptr<PdfBrush> brush = PdfBrushes::GetBlack();
		intrusive_ptr<PdfPen> pen = new PdfPen(brush, 0.75f);
		LPCWSTR_S arial = L"Arial";
		intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(arial, 12.f, PdfFontStyle::Italic, true);

		intrusive_ptr<PdfStringFormat> format = new PdfStringFormat(PdfTextAlignment::Right);
		format->SetMeasureTrailingSpaces(true);
		float space = font->GetHeight() * 0.75f;
		float x = margin->GetLeft();
		float width = page->GetCanvas()->GetClientSize()->GetWidth() - margin->GetLeft() - margin->GetRight();
		float y = page->GetCanvas()->GetClientSize()->GetHeight() - margin->GetBottom() + space;
		page->GetCanvas()->DrawLine(pen, x, y, x + width, y);
		y = y + 1;
		wstring numberLabel = formatSimple(L"{0} of {1}", startNumber++, pageCount);
		page->GetCanvas()->DrawString(numberLabel.c_str(), font, brush, x + width, y, format);
		page->GetCanvas()->SetTransparency(1);

	}
}

int main()
{
		wstring outputFile = OUTPUTPATH"PageNumberInFooter.pdf";
		//Load the Pdf from disk
		wstring inputFile = DATAPATH"MultipagePDF.pdf";
		intrusive_ptr<PdfDocument> doc = new PdfDocument();
		doc->LoadFromFile(inputFile.c_str());
		//Set the margin
		intrusive_ptr<PdfMargins> margin = doc->GetPageSettings()->GetMargins();
		//Draw Page number
		DrawPageNumber(doc, margin, 1, doc->GetPages()->GetCount());
		//Save the document
		doc->SaveToFile(outputFile.c_str());
		doc->Close();
}
