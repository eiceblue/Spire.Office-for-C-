#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile_1 = input_path + L"ReplaceContentWithDoc.docx";
	wstring inputFile_2 = input_path + L"Insert.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceContentWithDoc.docx";

	//Create the first document
	intrusive_ptr<Document> document1 = new Document();

	//Load the first document from disk.
	document1->LoadFromFile(inputFile_1.c_str());

	//Create the second document
	intrusive_ptr<Document>  document2 = new Document();

	//Load the second document from disk.
	document2->LoadFromFile(inputFile_2.c_str());

	//Get the first section of the first document 
	intrusive_ptr<Section> section1 = document1->GetSections()->GetItemInSectionCollection(0);

	//Create a regex
	intrusive_ptr<Regex> regex = new Regex(L"\\[MY_DOCUMENT\]", RegexOptions::None);

	//Find the text by regex
	vector<intrusive_ptr<TextSelection>> textSections = document1->FindAllPattern(regex);

	//Travel the found strings
	for (auto seletion : textSections)
	{
		intrusive_ptr<Paragraph> para = seletion->GetAsOneRange()->GetOwnerParagraph();

		intrusive_ptr<TextRange> textRange = seletion->GetAsOneRange();

		int index = section1->GetBody()->GetChildObjects()->IndexOf(para);

		for (int i = 0; i < document2->GetSections()->GetCount(); i++)
		{
			intrusive_ptr<Section> section2 = document2->GetSections()->GetItemInSectionCollection(i);
			for (int j = 0; j < section2->GetParagraphs()->GetCount(); j++)
			{
				intrusive_ptr<Paragraph> paragraph = section2->GetParagraphs()->GetItemInParagraphCollection(j);
				section1->GetBody()->GetChildObjects()->Insert(index, Object::Dynamic_cast<Paragraph>(paragraph->Clone()));
			}
		}

		para->GetChildObjects()->Remove(textRange);
	}

	document1->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document1->Dispose();
	document2->Dispose();
}