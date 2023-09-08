#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FormFieldTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddRadioButtonField.pdf";

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

	float x = 50;
	float y = 550;
	float tempX = 0;

	std::wstring text = L"RadioButton: ";

	//Draw a text into page
	page->GetCanvas()->DrawString(text.c_str(), font, brush, x, y);

	tempX = font->MeasureString(text.c_str())->GetWidth() + x + 15;

	//Create a pdf radio button field
	intrusive_ptr<PdfRadioButtonListField> radioButton = new PdfRadioButtonListField(page, L"RadioButton");
	radioButton->SetRequired(true);
	intrusive_ptr<PdfRadioButtonListItem> fieldItem = new PdfRadioButtonListItem();
	fieldItem->SetBorderWidth(0.75f);
	fieldItem->SetBounds(new RectangleF(tempX, y, 15, 15));
	radioButton->GetItems()->Add(fieldItem);

	//Add the radio button field into pdf document
	pdf->GetForm()->GetFields()->Add(radioButton);
	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}