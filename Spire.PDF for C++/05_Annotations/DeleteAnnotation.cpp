#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DeleteAnnotation.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DeleteAnnotation.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->GetPages()->GetItem(0)->GetAnnotationsWidget()->RemoveAt(0);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

