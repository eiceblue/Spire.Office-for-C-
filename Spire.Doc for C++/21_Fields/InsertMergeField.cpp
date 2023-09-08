#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertMergeField.docx";

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	intrusive_ptr<Paragraph> par = section->AddParagraph();

	//Add merge field in the paragraph
	intrusive_ptr<MergeField> field = Object::Dynamic_cast<MergeField>(par->AppendField(L"MyFieldName", FieldType::FieldMergeField));

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
