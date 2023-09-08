#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_1.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExpandBookmarks.pdf";
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	doc->GetViewerPreferences()->SetBookMarkExpandOrCollapse(true);

	doc->SaveToFile(outputFile.c_str());

	doc->Close();
}

