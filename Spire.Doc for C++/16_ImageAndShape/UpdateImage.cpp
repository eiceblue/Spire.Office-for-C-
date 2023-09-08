#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"ImageTemplate.docx";
	wstring imagePath = input_path + L"E-iceblue.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UpdateImage.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get all pictures in the Word document
	std::vector<intrusive_ptr<DocumentObject>> pictures;
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < sec->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> para = sec->GetParagraphs()->GetItemInParagraphCollection(j);
			for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
			{
				intrusive_ptr<DocumentObject> docObj = para->GetChildObjects()->GetItem(k);
				if (docObj->GetDocumentObjectType() == DocumentObjectType::Picture)
				{
					pictures.push_back(docObj);
				}
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
