#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"AllowBreakAcrossPages.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AllowBreakAcrossPages.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	for (int i = 0; i < table->GetRows()->GetCount(); i++)
	{
		intrusive_ptr<TableRow> row = table->GetRows()->GetItemInRowCollection(i);
		//Allow break across pages
		row->GetRowFormat()->SetIsBreakAcrossPages(true);
	}

	//Save the Word document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
