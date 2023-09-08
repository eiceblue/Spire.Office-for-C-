#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DeleteBookmark.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DeleteBookmark.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	doc->GetBookmarks()->RemoveAt(0);

	doc->SaveToFile(outputFile.c_str());

	doc->Close();
}


