#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_1.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddCheckBox.pdf";


	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->SetAllowCreateForm(true);
	intrusive_ptr<PdfCheckBoxField> checkboxField = new PdfCheckBoxField(doc->GetPages()->GetItem(0), L"fieldID");
	float checkboxWidth = 40;
	float checkboxHeight = 40;
	checkboxField->SetBounds(new RectangleF(60, 300, checkboxWidth, checkboxHeight));
	checkboxField->SetBorderWidth(0.75f);
	checkboxField->SetChecked(true);
	checkboxField->SetStyle(PdfCheckBoxStyle::Check);
	checkboxField->SetRequired(true);
	doc->GetForm()->GetFields()->Add(checkboxField);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}
