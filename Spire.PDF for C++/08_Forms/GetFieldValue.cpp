#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TextBoxSampleB_1.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetFieldValue.txt";

	//Create a pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	//Load from file
	doc->LoadFromFile(inputFile.c_str());

	//Get form fields
	intrusive_ptr<PdfFormWidget> formWidget = Object::Convert<PdfFormWidget>(doc->GetForm());

	//Get textbox
	intrusive_ptr<PdfTextBoxFieldWidget> textbox = Object::Convert<PdfTextBoxFieldWidget>(formWidget->GetFieldsWidget()->GetItem(L"Text1"));

	//Get the text of the textbox
	std::wstring text = textbox->GetText();

	//Save to file.
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << L"The value of the field named ";
	os << textbox->GetName();
	os << L" is " + text;
	os.close();

	doc->Close();
}
