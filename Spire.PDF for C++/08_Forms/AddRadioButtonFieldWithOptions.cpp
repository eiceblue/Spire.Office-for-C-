#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
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

template<typename T>
static wstring formatSimple(const wstring& input, T arg)
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
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FormFieldTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddRadioButtonFieldWithOptions.pdf";

	//Open pdf document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile(inputFile.c_str());

	//Get the first page
	intrusive_ptr<PdfPageBase> page = pdf->GetPages()->GetItem(0);

	//As for existing pdf, the property needs to be set as true
	pdf->SetAllowCreateForm(true);

	//Create a new pdf font
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 12.0f, PdfFontStyle::Bold);

	//Create a pdf brush
	intrusive_ptr<PdfBrush> brush = PdfBrushes::GetBlack();

	float x = 150;
	float y = 550;
	float temX = 0;
	//Create a pdf radio button list
	intrusive_ptr<PdfRadioButtonListField> radioButton = new PdfRadioButtonListField(page, L"RadioButton");
	radioButton->SetRequired(true);
	//Add items into radio button list.
	for (int i = 0; i < 3; i++)
	{
		intrusive_ptr<PdfRadioButtonListItem> item = new PdfRadioButtonListItem(formatSimple(L"item{0}", i).c_str());
		item->SetBorderWidth(0.75f);
		item->SetBounds(new RectangleF(x, y, 15, 15));
		item->SetBorderColor(new PdfRGBColor(Color::GetRed()));
		item->SetForeColor(new PdfRGBColor(Color::GetRed()));
		radioButton->GetItems()->Add(item);
		temX = x + 20;
		page->GetCanvas()->DrawString(formatSimple(L"Item{0}", i).c_str(), font, brush, temX, y);
		x = temX + 100;
	}

	//Add the radio button list field into pdf document
	pdf->GetForm()->GetFields()->Add(radioButton);
	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}