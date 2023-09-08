#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_1.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetBookmarkPageNumber.txt";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	intrusive_ptr<PdfBookmarkCollection> bookmarks = doc->GetBookmarks();
	intrusive_ptr<PdfBookmark> bookmark = bookmarks->GetItem(0);
	int pageNumber = doc->GetPages()->IndexOf(bookmark->GetDestination()->GetPage()) + 1;
	wstring showPageNumber = to_wstring(pageNumber);
	wstring text = L"The page number of the first bookmark is: " + showPageNumber;
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << text;
	os.close();

	doc->Close();
}


