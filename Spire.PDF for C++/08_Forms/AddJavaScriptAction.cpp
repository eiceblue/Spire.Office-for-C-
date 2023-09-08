#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;


int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FormFieldTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddJavaScriptAction.pdf";

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

	//Draw a text into page
	std::wstring text1 = L"Enter a number, such as 12345: ";
	//Draw a text into page
	page->GetCanvas()->DrawString(text1.c_str(), font, brush, x, y);

	//Add a textBox field 
	tempX = font->MeasureString(text1.c_str())->GetWidth() + x + 15;
	intrusive_ptr<PdfTextBoxField> textbox = new PdfTextBoxField(page, L"Number-TextBox");
	textbox->SetBounds(new RectangleF(tempX, y, 100, 15));
	textbox->SetBorderWidth(0.75f);
	textbox->SetBorderStyle(PdfBorderStyle::Solid);

	//Add a JavaScript action to be performed when uses type a keystroke into a text field
	wstring str = PdfJavaScript::GetSpecialKeystrokeString(3);
	std::wstring js = PdfJavaScript::GetNumberKeystrokeString(2, 0, 0, 0, L"$", true);
	intrusive_ptr<PdfJavaScriptAction> jsAction = new PdfJavaScriptAction(js.c_str());
	textbox->GetActions()->SetKeyPressed(jsAction);

	//Add a JavaScript action to format the value of text field
	js = PdfJavaScript::GetNumberFormatString(2, 0, 0, 0, L"$", true);
	jsAction = new PdfJavaScriptAction(js.c_str());
	textbox->GetActions()->SetFormat(jsAction);
	pdf->GetForm()->GetFields()->Add(textbox);

	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}
