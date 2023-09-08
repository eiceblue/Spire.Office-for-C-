#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	//Input and output file paths
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SetExportValueForCheckbox.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetExportValueForCheckbox.pdf";

	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	//Load from disk
	pdf->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfFormWidget> formWidget = Object::Convert<PdfFormWidget>(pdf->GetForm());
	int count = 1;
	//Traverse all FieldsWidget
	for (int i = 0; i < formWidget->GetFieldsWidget()->GetCount(); i++)
	{
		intrusive_ptr<PdfField> field = formWidget->GetFieldsWidget()->GetItem(i);
		//Find the checkbox
		wchar_t nm_w[100];
		swprintf(nm_w, 100, L"%hs", typeid(PdfCheckBoxWidgetFieldWidget).name());
		LPCWSTR_S newName = nm_w;
		if (wcscmp(newName, field->GetInstanceTypeName()) == 0)
		{
			intrusive_ptr<PdfCheckBoxWidgetFieldWidget> checkbox = Object::Convert<PdfCheckBoxWidgetFieldWidget>(field);
			//Set export value for checkbox
			checkbox->SetExportValue((L"True" + to_wstring(count++)).c_str());
		}
	}
	//Save the pdf file
	pdf->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	pdf->Close();
}

