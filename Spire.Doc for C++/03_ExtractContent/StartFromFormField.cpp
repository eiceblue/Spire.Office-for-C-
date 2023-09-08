#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TextInputField.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"StartFromFormField.docx";


	intrusive_ptr<Document> sourceDocument = new Document();

	
	sourceDocument->LoadFromFile(inputFile.c_str());


	intrusive_ptr<Document> destinationDoc = new Document();


	intrusive_ptr<Section> section = destinationDoc->AddSection();

	int index = 0;


	for (int i = 0; i < sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetCount(); i++)
	{
		intrusive_ptr<FormField> field = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetFormFields()->GetItem(i);

		if (field->GetType() == FieldType::FieldFormTextInput)
		{
			//Get the paragraph
			intrusive_ptr<Paragraph> paragraph = field->GetOwnerParagraph();

			//Get the index
			index = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->IndexOf(paragraph);
			break;
		}
	}

	//Extract the content
	for (int i = index; i < index + 3; i++)
	{
		//Clone the ChildObjects of source document
		intrusive_ptr<DocumentObject> doobj = sourceDocument->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->GetItem(i)->Clone();

		//Add to destination document 
		section->GetBody()->GetChildObjects()->Add(doobj);
	}

	//Save the document.
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	destinationDoc->Close();
}
