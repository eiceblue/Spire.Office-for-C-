#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetLocaleForField.docx";

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	intrusive_ptr<Paragraph> par = section->AddParagraph();

	//Add a date field
	intrusive_ptr<Field> field = par->AppendField(L"DocDate", FieldType::FieldDate);

	//Set the LocaleId for the textrange
	(Object::Dynamic_cast<TextRange>(field->GetOwnerParagraph()->GetChildObjects()->GetItem(0)))->GetCharacterFormat()->SetLocaleIdASCII(1049);

	field->SetFieldText(L"2019-10-10");
	//Update field
	document->SetIsUpdateFields(true);

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}