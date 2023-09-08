#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"ImageTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ResetImageSize.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first secion
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
	//Get the first paragraph
	intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(0);

	//Reset the image size of the first paragraph
	for (int i = 0; i < paragraph->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> docObj = paragraph->GetChildObjects()->GetItem(i);
		if (Object::Dynamic_cast<DocPicture>(docObj) != nullptr)
		{
			intrusive_ptr<DocPicture> picture = Object::Dynamic_cast<DocPicture>(docObj);
			picture->SetWidth(50.0f);
			picture->SetHeight(50.0f);
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
