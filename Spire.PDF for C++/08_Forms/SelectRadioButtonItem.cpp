#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RadioButton.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SelectRadioButtonItem.pdf";

	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile(inputFile.c_str());

	//Get pdf forms
	intrusive_ptr<PdfFormWidget> formWidget = Object::Convert<PdfFormWidget>(pdf->GetForm());

	//Find the radio button field and select the second item
	for (int i = 0; i < formWidget->GetFieldsWidget()->GetCount(); i++)
	{
		intrusive_ptr<PdfField> field = formWidget->GetFieldsWidget()->GetItem(i);

		wchar_t nm_w[100];
		swprintf(nm_w, 100, L"%hs", typeid(PdfRadioButtonListFieldWidget).name());
		LPCWSTR_S newName = nm_w;
		if (wcscmp(newName, field->GetInstanceTypeName()) == 0)
		{
			intrusive_ptr<PdfRadioButtonListFieldWidget> radioButton = Object::Convert<PdfRadioButtonListFieldWidget>(field);
			wstring str = radioButton->GetName();
			if (str == L"RadioButton")
			{
				radioButton->SetSelectedIndex(1);
			}
		}
	}
	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}
