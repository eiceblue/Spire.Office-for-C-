#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ExpandSpecificBookmarks.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExpandSpecificBookmarks.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	doc->GetBookmarks()->GetItem(0)->SetExpandBookmark(true);
	intrusive_ptr<PdfBookmarkCollection> bookMarkCo = Object::Convert<PdfBookmarkCollection>(doc->GetBookmarks()->GetItem(1));
	bookMarkCo->GetItem(0)->SetExpandBookmark(true);

	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}


