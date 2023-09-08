#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RemoveEditableRange.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveEditableRange.docx";

	//Create a new document
	intrusive_ptr<Document> document = new Document();
	//Load file from disk
	document->LoadFromFile(inputFile.c_str());
	//Find "PermissionStart" and "PermissionEnd" tags and remove them
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetBody()->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> para = section->GetBody()->GetParagraphs()->GetItemInParagraphCollection(j);
			for (int k = 0; k < para->GetChildObjects()->GetCount(); k++)
			{
				intrusive_ptr<DocumentObject> obj = para->GetChildObjects()->GetItem(k);
				if (Object::Dynamic_cast<PermissionStart>(obj) != nullptr || Object::Dynamic_cast<PermissionEnd>(obj) != nullptr)
				{
					para->GetChildObjects()->Remove(obj);
				}
				else
				{
					k++;
				}
			}
		}
	}


	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}