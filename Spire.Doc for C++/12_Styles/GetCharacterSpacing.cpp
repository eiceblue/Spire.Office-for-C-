#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Insert.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetCharacterSpacing.txt";

	//Create a document
	intrusive_ptr<Document> document = new Document();

	//Load the document from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the first section of document
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	//Get the first paragraph 
	intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(0);

	//Define two variables
	std::wstring fontName = L"";
	float fontSpacing = 0;

	//Traverse the ChildObjects 
	for (int i = 0; i < paragraph->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> docObj = paragraph->GetChildObjects()->GetItem(i);
		//If it is TextRange
		if (Object::Dynamic_cast<TextRange>(docObj) != nullptr)
		{
			intrusive_ptr<TextRange> textRange = Object::Dynamic_cast<TextRange>(docObj);

			//Get the character spacing
			fontSpacing = textRange->GetCharacterFormat()->GetCharacterSpacing();
		}
	}
	std::wstring content;
	content.append(L"The font of first paragraph is ")
		.append(fontName)
		.append(L", the character spacing is ")
		.append(std::to_wstring(fontSpacing))
		.append(L"pt.");
	std::wofstream foo(outputFile);
	foo << content;
	foo.close();
	document->Close();
}

