#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Bookmark.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FromBookmark.docx";

	//Create the source document
	intrusive_ptr<Document> sourcedocument = new Document();

	//Load the source document from disk.
	sourcedocument->LoadFromFile(inputFile.c_str());

	//Create a destination document
	intrusive_ptr<Document> destinationDoc = new Document();

	//Add a section for destination document
	intrusive_ptr<Section> section = destinationDoc->AddSection();

	//Add a paragraph for destination document
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Locate the bookmark in source document
	intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(sourcedocument);

	//Find bookmark by name
	navigator->MoveToBookmark(L"Test", true, true);

	//Get text body part
	intrusive_ptr<TextBodyPart> textBodyPart = navigator->GetBookmarkContent();

	//Create a textRange type list
	std::vector<intrusive_ptr<TextRange>> list;

	//Traverse the items of text Body
	for (int i = 0; i < textBodyPart->GetBodyItems()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> item = textBodyPart->GetBodyItems()->GetItem(i);
		//if it is paragraph
		if (Object::Dynamic_cast<Paragraph>(item) != nullptr)
		{
			//Traverse the ChildObjects of the paragraph
			for (int i = 0; i < (Object::Dynamic_cast<Paragraph>(item))->GetChildObjects()->GetCount(); i++)
			{
				intrusive_ptr<DocumentObject> childObject = (Object::Dynamic_cast<Paragraph>(item))->GetChildObjects()->GetItem(i);
				//if it is TextRange
				if (Object::Dynamic_cast<TextRange>(childObject) != nullptr)
				{
					//Add it into list
					intrusive_ptr<TextRange> range = Object::Dynamic_cast<TextRange>(childObject);
					list.push_back(range);
				}
			}
		}
	}

	//Add the extracted content to destination document
	for (int m = 0; m < list.size(); m++)
	{
		paragraph->GetChildObjects()->Add(list[m]->Clone());
	}

	//Save the document.
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	destinationDoc->Close();
}