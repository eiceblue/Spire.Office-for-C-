#include "../pch.h"

using namespace Spire::Pdf;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path+L"PDFTemplate-Az.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path +L"ExtractTextFromParticularPage.txt";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	// Extract text from page keeping white space
	wstring text = page->ExtractText(true);
	// Extract text from page without keeping white space
	//wstring text = page->ExtractText(false);
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << text;
	os.close();
	doc->Close();
}

