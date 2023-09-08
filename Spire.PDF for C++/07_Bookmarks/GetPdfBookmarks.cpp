#include "../pch.h"
#include"../stringbuilder.h"

using namespace Spire::Pdf;
using namespace std;

const wstring ToString(PdfTextStyle style) {
	switch (style)
	{
	case Spire::Pdf::PdfTextStyle::Regular:
		return L"Regular";
		break;
	case Spire::Pdf::PdfTextStyle::Italic:
		return L"Italic";
		break;
	case Spire::Pdf::PdfTextStyle::Bold:
		return L"Bold";
		break;
	default:
		break;
	}
}


void GetChildBookmark(PdfBookmark* parentBookmark, StringBuilder* content)
{
	if (parentBookmark->GetCount() > 0)
	{
		for (int i = 0; i < parentBookmark->GetCount(); i++)
		{
			//Get the title.
			PdfBookmark* childBookmark = parentBookmark->GetItem(i);
			content->appendLine(childBookmark->GetTitle());

			//Get the text style.
			wstring textStyle = ToString(childBookmark->GetDisplayStyle());
			content->appendLine(textStyle);
			GetChildBookmark(childBookmark, content);
		}
	}
}


void GetBookmarks(PdfBookmarkCollection *bookmarks, const std::wstring &result)
{
	//Declare a new StringBuilder content
	StringBuilder *content = new StringBuilder();

	//Get Pdf bookmarks information.
	if (bookmarks->GetCount()> 0)
	{
		content->appendLine(L"Pdf bookmarks:");
		for (int i =0; i < bookmarks->GetCount();i++)
		{
			PdfBookmark *parentBookmark = bookmarks->GetItem(i);
			//Get the title.
			content->appendLine(parentBookmark->GetTitle());

			//Get the text style.
			wstring textStyle = ToString(parentBookmark->GetDisplayStyle());
			content->appendLine(textStyle);
			GetChildBookmark(parentBookmark,content);
		}
	}
	//Save to file.
	wofstream os;
	os.open(result, ios::trunc);
}

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_1.pdf)";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetPdfBookmarks.txt";
	//Create a new Pdf document.
	PdfDocument* doc = new PdfDocument();

	//Load the file from disk.
	doc->LoadFromFile(inputFile.c_str());

	//Get bookmarks collection of the Pdf file.
	PdfBookmarkCollection* bookmarks = doc->GetBookmarks();;

	//Get Pdf Bookmarks.
	GetBookmarks(bookmarks, outputFile);

	delete doc;
}