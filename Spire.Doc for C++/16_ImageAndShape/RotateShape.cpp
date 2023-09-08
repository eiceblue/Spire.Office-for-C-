#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Shapes.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RotateShape.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Traverse the word document and set the shape rotation as 20
	for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
	{
		intrusive_ptr<Paragraph> para = section->GetParagraphs()->GetItemInParagraphCollection(i);
		for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(j);
			if (Object::Dynamic_cast<ShapeObject>(obj) != nullptr)
			{
				(Object::Dynamic_cast<ShapeObject>(obj))->SetRotation(20.0);
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
