#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"PageRef.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertPageRefField.docx";

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetLastSection();

	intrusive_ptr<Paragraph> par = section->AddParagraph();

	//Add page ref field
	intrusive_ptr<Field> field = par->AppendField(L"pageRef", FieldType::FieldPageRef);

	//Set field code
	field->SetCode(L"PAGEREF  bookmark1 \\# \"0\" \\* Arabic  \\* MERGEFORMAT");

	//Update field
	document->SetIsUpdateFields(true);

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
