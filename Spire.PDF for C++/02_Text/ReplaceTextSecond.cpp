#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path+L"ReplaceTextInPage.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"ReplaceTextSecond.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfTextReplacer> replacer = new PdfTextReplacer(page);
	replacer->ReplaceAllText(L"Spire.PDF", L"E-iceblue");
	replacer->ReplaceAllText(L"Adobe Acrobat", L"PDF editors");
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

