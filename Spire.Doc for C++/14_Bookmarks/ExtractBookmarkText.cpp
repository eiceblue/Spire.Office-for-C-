#include "pch.h"
#include <locale>
#include <codecvt>
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"BookmarkTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExtractBookmarkText.txt";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Creates a BookmarkNavigator instance to access the bookmark
	intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(doc);
	//Locate a specific bookmark by bookmark name
	navigator->MoveToBookmark(L"Content");
	intrusive_ptr<TextBodyPart> textBodyPart = navigator->GetBookmarkContent();

	//Iterate through the items in the bookmark content to get the text
	std::wstring text = L"";
	for (int i = 0; i < textBodyPart->GetBodyItems()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> item = textBodyPart->GetBodyItems()->GetItem(i);
		if (Object::Dynamic_cast<Paragraph>(item) != nullptr)
		{
			intrusive_ptr<Paragraph> paragraph = (Object::Dynamic_cast<Paragraph>(item));
			for (int j = 0; j < paragraph->GetChildObjects()->GetCount(); j++)
			{
				intrusive_ptr<DocumentObject> childObject = paragraph->GetChildObjects()->GetItem(j);
				if (Object::Dynamic_cast<TextRange>(childObject) != nullptr)
				{
					text += (Object::Dynamic_cast<TextRange>(childObject))->GetText();
				}
			}
		}
	}

	//Save to TXT File and launch it
	std::wofstream foo(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	foo.imbue(LocUtf8);
	foo << text;
	foo.close();
	doc->Close();
}
