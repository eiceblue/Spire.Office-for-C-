#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ReplaceTextInTable.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetRowCellIndex.txt";

	//Load Word from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);

	//Get the first table in the section
	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(section->GetTables()->GetItemInTableCollection(0));

	wstring* content = new wstring();

	//Get table collections
	intrusive_ptr<TableCollection> collections = section->GetTables();

	//Get the table index
	int tableIndex = collections->IndexOf(table);

	//Get the index of the last table row
	intrusive_ptr<TableRow> row = table->GetLastRow();
	int rowIndex = row->GetRowIndex();

	//Get the index of the last table cell
	intrusive_ptr<TableCell> cell = Object::Dynamic_cast<TableCell>(row->GetLastChild());
	int cellIndex = cell->GetCellIndex();

	//Append these information into content
	content->append(L"Table index is " + std::to_wstring(tableIndex) + L"\r\n");
	content->append(L"Row index is " + std::to_wstring(rowIndex) + L"\r\n");
	content->append(L"Cell index is " + std::to_wstring(cellIndex) + L"\r\n");
	//Save to txt file
	wofstream write(outputFile);
	write << content->c_str();
	write.close();
	doc->Close();
	delete content;
}
