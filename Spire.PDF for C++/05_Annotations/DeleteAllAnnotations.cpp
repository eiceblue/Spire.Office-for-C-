#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {

	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_3.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DeleteAllAnnotations.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->GetPages()->GetItem(0)->GetAnnotationsWidget()->Clear();

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

