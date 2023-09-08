#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Bookmarks.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetBookmarks.txt";

	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the bookmark by index.
	intrusive_ptr<Bookmark> bookmark1 = document->GetBookmarks()->GetItem(0);

	//Get the bookmark by name.
	intrusive_ptr<Bookmark> bookmark2 = document->GetBookmarks()->GetItem(L"Test2");

	//Create StringBuilder to save 
	wstring* content = new wstring();

	//Set string format for displaying
	content->append(L"The bookmark obtained by index is ");
	content->append(bookmark1->GetName());
	content->append(L".\n");
	content->append(L"The bookmark obtained by name is ");
	content->append(bookmark2->GetName());
	content->append(L".\n");
	
	//Save them to a txt file
	std::wofstream foo(outputFile);
	foo << content->c_str();
	foo.close();
	document->Close();
	delete content;
}
