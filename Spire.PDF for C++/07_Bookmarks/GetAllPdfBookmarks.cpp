#include "../pch.h"

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

void GetChildBookmark(intrusive_ptr<PdfBookmark> parentBookmark, wstring& content)
{
	if (parentBookmark->GetCount() > 0)
	{
		for (int i = 0; i < parentBookmark->GetCount(); i++)
		{
			//Get the title.
			content += parentBookmark->GetItem(i)->GetTitle();
			//Get the text style.
			wstring textStyle = ToString(parentBookmark->GetItem(i)->GetDisplayStyle());
			content += L"\n" + textStyle + L"\n";
			GetChildBookmark(parentBookmark->GetItem(i), content);
		}
	}

}
		
		

void GetBookmarks(intrusive_ptr<PdfBookmarkCollection> bookmarks, std::wstring result)
{
	wstring content;

	//Get Pdf bookmarks information.
	if (bookmarks->GetCount() > 0)
	{
		content += L"Pdf bookmarks: \n";
		for (int i = 0; i < bookmarks->GetCount(); i++)
		{
			content += bookmarks->GetItem(i)->GetTitle();
			content += L"\n";
			//Get the text style.
			wstring textStyle = ToString(bookmarks->GetItem(i)->GetDisplayStyle());
			content += textStyle + L"\n";

			GetChildBookmark(bookmarks->GetItem(i), content);
		}
	}

	//Save to file.
	wofstream os;
	os.open(result, ios::trunc);
	os << content;
	os.close();

}

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_1.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetAllPdfBookmarks.txt";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	intrusive_ptr<PdfBookmarkCollection> bookmarks = doc->GetBookmarks();
	GetBookmarks(bookmarks, outputFile);

	doc->Close();
}

