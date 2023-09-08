#include "pch.h"
using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddAlternativeText.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Get the first table in the section
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	//Add alternative text
	//Add title
	table->SetTitle(L"Table 1");
	//Add description
	table->SetTableDescription(L"Description Text");

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
