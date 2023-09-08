#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_1.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetInheritZoom.pdf";


	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	intrusive_ptr<PdfBookmarkCollection> bookmarks = doc->GetBookmarks();
	//Set Zoom level as 0, which the value is inherit zoom

	for (int i = 0; i < doc->GetBookmarks()->GetCount(); i++)
	{
		doc->GetBookmarks()->GetItem(i)->GetDestination()->SetZoom(0.5f);
	}
	doc->SaveToFile(outputFile.c_str());

	doc->Close();
}

