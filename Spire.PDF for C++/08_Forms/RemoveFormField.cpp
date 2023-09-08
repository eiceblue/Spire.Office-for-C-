#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RemoveFormField.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveFormField.pdf";

	//Create a PdfDocument
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	//Load the input file from disk
	pdf->LoadFromFile(inputFile.c_str());
	//Get form from the document
	intrusive_ptr<PdfFormWidget> formWidget = Object::Convert<PdfFormWidget>(pdf->GetForm());
	if (formWidget != nullptr)
	{
		for (int i = 0; i <= formWidget->GetFieldsWidget()->GetCount() - 1; i++)
		{
			//Case 1: Remove the first form field
			if (i == 0)
			{
				intrusive_ptr<PdfField> field = formWidget->GetFieldsWidget()->GetItem(i);
				formWidget->GetFieldsWidget()->Remove(field);
				break;
			}
		}
		//Case 2: Remove all form fields
		//formWidget.FieldsWidget.Clear();

		//Save the pdf file
		pdf->SaveToFile(outputFile.c_str());
		pdf->Close();
	}
}
