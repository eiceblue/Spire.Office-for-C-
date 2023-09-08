#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"ImageTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetTransparentColorForImage.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first paragraph in the first section
	intrusive_ptr<Paragraph> paragraph = doc->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0);

	//Set the blue color of the image(s) in the paragraph to transperant
	for (int i = 0; i < paragraph->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> obj = paragraph->GetChildObjects()->GetItem(i);
		if (Object::Dynamic_cast<DocPicture>(obj) != nullptr)
		{
			intrusive_ptr<DocPicture> picture = Object::Dynamic_cast<DocPicture>(obj);
			picture->SetTransparentColor(Color::GetBlue());
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
