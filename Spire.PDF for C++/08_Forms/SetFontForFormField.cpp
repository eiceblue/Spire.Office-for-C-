#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TextBoxSampleB.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetFontForFormField.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Get form fields
	intrusive_ptr<PdfFormWidget> formWidget = Object::Convert<PdfFormWidget>(doc->GetForm());

	//Get textbox
	intrusive_ptr<PdfTextBoxFieldWidget> textbox = Object::Convert<PdfTextBoxFieldWidget>(formWidget->GetFieldsWidget()->GetItem(L"Text1"));

	//Set the font for textbox
	textbox->SetFont(new PdfTrueTypeFont(L"Tahoma", 12, PdfFontStyle::Regular, true));
	//Set text
	textbox->SetText(L"Hello World");
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
