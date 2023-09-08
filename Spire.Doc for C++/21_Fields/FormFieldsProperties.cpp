#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FillFormField.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FormFieldsProperties.docx";

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	//Get FormField by index
	intrusive_ptr<FormField> formField = section->GetBody()->GetFormFields()->GetItem(1);

	if (formField->GetType() == FieldType::FieldFormTextInput)
	{
		wstring formFieldName = formField->GetName();
		wstring temp = L"My name is " + formFieldName;
		formField->SetText(temp.c_str());
		formField->GetCharacterFormat()->SetTextColor(Color::GetRed());
		formField->GetCharacterFormat()->SetItalic(true);
	}

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
