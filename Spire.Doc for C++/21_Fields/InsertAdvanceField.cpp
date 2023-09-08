#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertAdvanceField.docx";

	//Open a Word document.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	intrusive_ptr<Paragraph> par = section->AddParagraph();

	//Add advance field
	intrusive_ptr<Field> field = par->AppendField(L"Field", FieldType::FieldAdvance);

	//Add field code
	field->SetCode(L"ADVANCE \\d 10 \\l 10 \\r 10 \\u 0 \\x 100 \\y 100 ");

	//Update field
	document->SetIsUpdateFields(true);

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}

