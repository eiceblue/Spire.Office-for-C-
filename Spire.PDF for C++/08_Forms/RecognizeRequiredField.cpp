#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RecognizeRequiredField.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RecognizeRequiredField.txt";


	//Open pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	//Get pdf forms
	intrusive_ptr<PdfFormWidget> formWidget = Object::Convert<PdfFormWidget>(doc->GetForm());
	wstring str = L"";
	for (int i = 0; i < formWidget->GetFieldsWidget()->GetCount(); i++)
	{
		intrusive_ptr<PdfField> field = formWidget->GetFieldsWidget()->GetItem(i);
		//Judge if the field is required
		if (field->GetRequired())
		{
			str.append(L"The field named: ");
			str.append(field->GetName());
			str.append(L" is required\r\n");
		}
	}
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << str;
	os.close();
	//Save the document
	doc->Close();
}