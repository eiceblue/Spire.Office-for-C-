#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TextBoxSampleB.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetCoordinates.txt";


	//Create a pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	//Load from file
	doc->LoadFromFile(inputFile.c_str());

	//Get form fields
	intrusive_ptr<PdfFormWidget> formWidget = Object::Convert<PdfFormWidget>(doc->GetForm());

	//Get textbox
	intrusive_ptr<PdfTextBoxFieldWidget> textbox = Object::Convert<PdfTextBoxFieldWidget>(formWidget->GetFieldsWidget()->GetItem(L"Text1"));

	//Get the location of the textbox
	intrusive_ptr<PointF> location = textbox->GetLocation();
	//Save to file.
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << L"The location of the field named ";
	os << textbox->GetName();
	os << L" is\n X:";
	os << to_wstring(location->GetX());
	os << L"  Y:";
	os << to_wstring(location->GetY());
	os.close();

	doc->Close();
}
