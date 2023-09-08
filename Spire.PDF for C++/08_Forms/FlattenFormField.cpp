#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FlattenFormField.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FlattenFormField.pdf";

	//Open pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	//Flatten form fields
	doc->GetForm()->SetIsFlatten(true);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
