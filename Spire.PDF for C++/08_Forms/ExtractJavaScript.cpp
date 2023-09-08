#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ExtractJavaScript.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExtractJavaScript.txt";

	//Create a pdf document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();

	//Load a pdf document
	pdf->LoadFromFile(inputFile.c_str());

	std::wstring js = L"";
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
			if (str == L"total")
			{
				//Get the action
				intrusive_ptr<PdfJavaScriptAction> jsa = textbox->GetActions()->GetCalculate();

				if (jsa != nullptr)
				{
					//Get JavaScript
					js = jsa->GetScript();
				}
			}
		}
	}

	//Save and launch the result file
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << js;
	os.close();
	pdf->Close();
}
