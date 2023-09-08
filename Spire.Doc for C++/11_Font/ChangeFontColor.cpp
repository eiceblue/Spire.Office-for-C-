#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Sample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ChangeFontColor.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section and first paragraph
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	intrusive_ptr<Paragraph> p1 = section->GetParagraphs()->GetItemInParagraphCollection(0);
	//Iterate through the childObjects of the paragraph 1 
	int p1ChildObjectsCount = p1->GetChildObjects()->GetCount();
	for (int i = 0; i < p1ChildObjectsCount; i++)
	{
		intrusive_ptr<DocumentObject> childObj = p1->GetChildObjects()->GetItem(i);
		if (Object::Dynamic_cast<TextRange>(childObj) != nullptr)
		{
			//Change text color
			intrusive_ptr<TextRange> tr = Object::Dynamic_cast<TextRange>(childObj);
			tr->GetCharacterFormat()->SetTextColor(Color::GetRosyBrown());
		}
	}

	//Get the second paragraph
	intrusive_ptr<Paragraph> p2 = section->GetParagraphs()->GetItemInParagraphCollection(1);

	//Iterate through the childObjects of the paragraph 2
	int p2ChildObjectsCount = p2->GetChildObjects()->GetCount();
	for (int i = 0; i < p2ChildObjectsCount; i++)
	{
		intrusive_ptr<DocumentObject> childObj = p2->GetChildObjects()->GetItem(i);
		if (Object::Dynamic_cast<TextRange>(childObj) != nullptr)
		{
			//Change text color
			intrusive_ptr<TextRange> tr = Object::Dynamic_cast<TextRange>(childObj);
			tr->GetCharacterFormat()->SetTextColor(Color::GetDarkGreen());
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
