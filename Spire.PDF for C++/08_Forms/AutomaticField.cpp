#include "../pch.h"

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


void DrawAutomaticField(const std::wstring& fieldName, intrusive_ptr<PdfPageBase> page, intrusive_ptr<RectangleF> bounds)
{
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.0f, PdfFontStyle::Italic, true);
	intrusive_ptr<PdfBrush> brush = PdfBrushes::GetOrangeRed();
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat(PdfTextAlignment::Left, PdfVerticalAlignment::Middle);

	if (L"DateTimeField" == fieldName)
	{
		intrusive_ptr<PdfDateTimeField> field = new PdfDateTimeField();
		field->SetFont(font);
		field->SetBrush(brush);
		field->SetStringFormat(format);
		field->SetBounds(bounds);
		field->SetDateFormatString(L"yyyy-MM-dd HH:mm:ss");
		Object::Dynamic_cast<PdfGraphicsWidget>(field)->Draw(page->GetCanvas());
	}

	if (L"CreationDateField" == fieldName)
	{
		intrusive_ptr<PdfCreationDateField> field = new PdfCreationDateField();
		field->SetFont(font);
		field->SetBrush(brush);
		field->SetStringFormat(format);
		field->SetBounds(bounds);
		field->SetDateFormatString(L"yyyy-MM-dd HH:mm:ss");
		Object::Dynamic_cast<PdfGraphicsWidget>(field)->Draw(page->GetCanvas());
	}

	if (L"DocumentAuthorField" == fieldName)
	{
		intrusive_ptr<PdfDocumentAuthorField> field = new PdfDocumentAuthorField();
		field->SetFont(font);
		field->SetBrush(brush);
		field->SetStringFormat(format);
		field->SetBounds(bounds);
		Object::Dynamic_cast<PdfGraphicsWidget>(field)->Draw(page->GetCanvas());
	}


	if (L"SectionNumberField" == fieldName)
	{
		intrusive_ptr<PdfSectionNumberField> field = new PdfSectionNumberField();
		field->SetFont(font);
		field->SetBrush(brush);
		field->SetStringFormat(format);
		field->SetBounds(bounds);
		Object::Dynamic_cast<PdfGraphicsWidget>(field)->Draw(page->GetCanvas());
	}

	if (L"SectionPageNumberField" == fieldName)
	{
		intrusive_ptr<PdfSectionPageNumberField> field = new PdfSectionPageNumberField();
		field->SetFont(font);
		field->SetBrush(brush);
		field->SetStringFormat(format);
		field->SetBounds(bounds);
		Object::Dynamic_cast<PdfGraphicsWidget>(field)->Draw(page->GetCanvas());
	}

	if (L"SectionPageCountField" == fieldName)
	{
		intrusive_ptr<PdfSectionPageCountField> field = new PdfSectionPageCountField();
		field->SetFont(font);
		field->SetBrush(brush);
		field->SetStringFormat(format);
		field->SetBounds(bounds);
		Object::Dynamic_cast<PdfGraphicsWidget>(field)->Draw(page->GetCanvas());
	}

	if (L"PageNumberField" == fieldName)
	{
		intrusive_ptr<PdfPageNumberField> field = new PdfPageNumberField();
		field->SetFont(font);
		field->SetBrush(brush);
		field->SetStringFormat(format);
		field->SetBounds(bounds);
		Object::Dynamic_cast<PdfGraphicsWidget>(field)->Draw(page->GetCanvas());
	}

	if (L"PageCountField" == fieldName)
	{
		intrusive_ptr<PdfPageCountField> field = new PdfPageCountField();
		field->SetFont(font);
		field->SetBrush(brush);
		field->SetStringFormat(format);
		field->SetBounds(bounds);
		Object::Dynamic_cast<PdfGraphicsWidget>(field)->Draw(page->GetCanvas());
	}

	if (L"DestinationPageNumberField" == fieldName)
	{
		intrusive_ptr<PdfDestinationPageNumberField> field = new PdfDestinationPageNumberField();
		field->SetFont(font);
		field->SetBrush(brush);
		field->SetStringFormat(format);
		field->SetBounds(bounds);
		field->SetPage(Object::Convert<PdfNewPage>(page));
		Object::Dynamic_cast<PdfGraphicsWidget>(field)->Draw(page->GetCanvas());
	}

	if (L"CompositeField" == fieldName)
	{
		intrusive_ptr<PdfSectionPageNumberField> field1 = new PdfSectionPageNumberField();
		field1->SetNumberStyle(PdfNumberStyle::LowerRoman);
		intrusive_ptr<PdfSectionPageCountField> field2 = new PdfSectionPageCountField();
		intrusive_ptr<PdfCompositeField> fields = new PdfCompositeField();
		fields->SetFont(font);
		fields->SetBrush(brush);
		fields->SetStringFormat(format);
		fields->SetBounds(bounds);
		fields->SetAutomaticFields({ field1, field2 });
		fields->SetText(L"section page {0} of {1}");
		Object::Dynamic_cast<PdfGraphicsWidget>(fields)->Draw(page->GetCanvas());
	}

}
void DrawAutomaticField(intrusive_ptr<PdfPageBase> page)
{
	float y = 20;

	//Title
	intrusive_ptr<PdfBrush> brush1 = PdfBrushes::GetCadetBlue();

	intrusive_ptr<PdfTrueTypeFont> font1 = new PdfTrueTypeFont(L"Arial", 16.0f, PdfFontStyle::Bold, true);

	intrusive_ptr<PdfStringFormat> format1 = new PdfStringFormat(PdfTextAlignment::Center);
	page->GetCanvas()->DrawString(L"Automatic Field List", font1, brush1, page->GetCanvas()->GetClientSize()->GetWidth() / 2, y, format1);
	y = y + font1->MeasureString(L"Automatic Field List", format1)->GetHeight();
	y = y + 15;

	std::vector<std::wstring> fieldList = { L"DateTimeField", L"CreationDateField", L"DocumentAuthorField", L"SectionNumberField", L"SectionPageNumberField", L"SectionPageCountField", L"PageNumberField", L"PageCountField", L"DestinationPageNumberField", L"CompositeField" };


	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.0f, PdfFontStyle::Regular, true);

	intrusive_ptr<PdfStringFormat> fieldNameFormat = new PdfStringFormat();
	fieldNameFormat->SetMeasureTrailingSpaces(true);
	for (auto fieldName : fieldList)
	{
		//Draw field name
		std::wstring text = formatSimple(L"{0}: ", fieldName);
		page->GetCanvas()->DrawString(text.c_str(), font, PdfBrushes::GetDodgerBlue(), 0, y);
		float x = font->MeasureString(text.c_str(), fieldNameFormat)->GetWidth();
		intrusive_ptr<RectangleF> bounds = new RectangleF(x, y, 200, font->GetHeight());
		DrawAutomaticField(fieldName, page, bounds);
		y = y + font->GetHeight() + 8;
	}
}

int main()
{
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AutomaticField.pdf";

	//Create a pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->GetDocumentInformation()->SetAuthor(L"Spire.Pdf");

	//Set the margin
	intrusive_ptr<PdfUnitConvertor> unitCvtr = new PdfUnitConvertor();
	intrusive_ptr<PdfMargins> margin = new PdfMargins();
	margin->SetTop(unitCvtr->ConvertUnits(2.54f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetBottom(margin->GetTop());
	margin->SetLeft(unitCvtr->ConvertUnits(3.17f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetRight(margin->GetLeft());

	//Create section
	intrusive_ptr<PdfSection> section = doc->GetSections()->Add();
	section->GetPageSettings()->SetSize(PdfPageSize::A4());
	section->GetPageSettings()->SetMargins(margin);

	//Create one page
	intrusive_ptr<PdfPageBase> page = section->GetPages()->Add();

	//Draw automatic fields
	DrawAutomaticField(page);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

		
