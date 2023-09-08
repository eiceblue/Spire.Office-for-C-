#include "../pch.h"

using namespace Spire::Pdf;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile =input_path+L"ExtractTextFromSpecificArea.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"ExtractTextFromSpecificArea.txt";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<RectangleF> rec = new RectangleF(80, 180, 500, 200);
	wstring text = page->ExtractText(rec);
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << text;
	os.close();
	doc->Close();
}

