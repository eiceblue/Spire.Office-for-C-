#include "pch.h"
using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddOrDeleteRow.docx";

	//Create a document
	intrusive_ptr<Document> document = new Document();
	//Load file
	document->LoadFromFile(inputFile.c_str());
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	//Delete the seventh row
	table->GetRows()->RemoveAt(7);

	//Add a row and insert it into specific position
	intrusive_ptr<TableRow> row = new TableRow(document);
	for (int i = 0; i < table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetCount(); i++)
	{
		intrusive_ptr<TableCell> tc = row->AddCell();
		intrusive_ptr<Paragraph> paragraph = tc->AddParagraph();
		paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
		paragraph->AppendText(L"Added");
	}
	table->GetRows()->Insert(2, row);
	//Add a row at the end of table
	table->AddRow();

	//Save to file and launch it
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
