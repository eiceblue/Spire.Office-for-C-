#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;


void EditChild2Bookmark(intrusive_ptr<PdfBookmark> childBookmark)
{
	for (int i = 0; i < childBookmark->GetCount(); i++)
	{
		intrusive_ptr<PdfBookmark> childbookmark = childBookmark->GetItem(i);
		childbookmark->SetColor(new PdfRGBColor(Spire::Common::Color::GetLightSalmon()));
		childbookmark->SetDisplayStyle(PdfTextStyle::Italic);
	}
}

void EditChildBookmark(intrusive_ptr<PdfBookmark> parentBookmark)
{
	for (int i = 0; i < parentBookmark->GetCount(); i++)
	{
		intrusive_ptr<PdfBookmark> childbookmark = parentBookmark->GetItem(i);
		childbookmark->SetColor(new PdfRGBColor(Spire::Common::Color::GetBlue()));
		childbookmark->SetDisplayStyle(PdfTextStyle::Regular);
		EditChild2Bookmark(childbookmark);
	}
}

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"UpdateBookmark.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UpdateBookmark.pdf";


	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first bookmark
	intrusive_ptr<PdfBookmark> bookMark = doc->GetBookmarks()->GetItem(0);

	//Change the title of the bookmark
	bookMark->SetTitle(L"Modified BookMark");

	//Set the color of the bookmark
	bookMark->SetColor(new PdfRGBColor(Spire::Common::Color::GetBlack()));

	//Set the outline text style of the bookmark
	bookMark->SetDisplayStyle(PdfTextStyle::Bold);

	//Edit child bookmarks of the parent bookmark
	EditChildBookmark(bookMark);

	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

