#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FormField.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ModifyFormFieldValue.pdf";

	//Open pdf document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile(inputFile.c_str());

	intrusive_ptr<PdfFormWidget> form = Object::Convert<PdfFormWidget>(pdf->GetForm());
	for (int i = 0; i < form->GetFieldsWidget()->GetCount(); i++)
	{
		intrusive_ptr<PdfField> field = form->GetFieldsWidget()->GetItem(i);
		wchar_t nm_w[100];
		swprintf(nm_w, 100, L"%hs", typeid(PdfTextBoxFieldWidget).name());
		LPCWSTR_S newName = nm_w;
		if (wcscmp(newName, field->GetInstanceTypeName()) == 0)
		{
			intrusive_ptr<PdfTextBoxFieldWidget> textbox = Object::Convert<PdfTextBoxFieldWidget>(field);

			//Find the textbox named total
			wstring str = textbox->GetName();
			if (str == L"TextBox1")
			{
				textbox->SetText(L"New value");
			}
		}
	}

	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();

}
