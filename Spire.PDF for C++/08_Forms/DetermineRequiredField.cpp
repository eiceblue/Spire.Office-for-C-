#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DetermineRequiredField.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DetermineRequiredField.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Get pdf forms
	intrusive_ptr<PdfFormWidget> formWidget = Object::Convert<PdfFormWidget>(doc->GetForm());

	//Find the particular form field and determine if it marks as required or not
	for (int i = 0; i < formWidget->GetFieldsWidget()->GetCount(); i++)
	{
		intrusive_ptr<PdfField> field = formWidget->GetFieldsWidget()->GetItem(i);
		wchar_t nm_w[100];
		swprintf(nm_w, 100, L"%hs", typeid(PdfTextBoxFieldWidget).name());
		LPCWSTR_S newName = nm_w;
		if (wcscmp(newName, field->GetInstanceTypeName()) == 0)
		{
			intrusive_ptr<PdfTextBoxFieldWidget> textbox = Object::Convert<PdfTextBoxFieldWidget>(field);
			wstring str = textbox->GetName();
			if (str == L"username")
			{
				textbox->SetRequired(true);
			}
			if (str == L"password2")
			{
				textbox->SetRequired(false);
			}
		}
	}
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
