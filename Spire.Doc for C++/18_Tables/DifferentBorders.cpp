#include "pch.h"
using namespace Spire::Doc;


void setTableBorders(intrusive_ptr<Table> table)
{
	table->GetTableFormat()->GetBorders()->SetBorderType(BorderStyle::Single);
	table->GetTableFormat()->GetBorders()->SetLineWidth(3.0F);
	table->GetTableFormat()->GetBorders()->SetColor(Color::GetRed());
}

void setCellBorders(intrusive_ptr<TableCell> tableCell)
{
	tableCell->GetCellFormat()->GetBorders()->SetBorderType(BorderStyle::DotDash);
	tableCell->GetCellFormat()->GetBorders()->SetLineWidth(1.0F);
	tableCell->GetCellFormat()->GetBorders()->SetColor(Color::GetGreen());
}
int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DifferentBorders.docx";

	//Open a Word document as template
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Table> table = Object::Dynamic_cast<Table>(document->GetSections()->GetItemInSectionCollection(0)->GetTables()->GetItemInTableCollection(0));

	//Set borders of table
	setTableBorders(table);

	//Set borders of cell
	setCellBorders(table->GetRows()->GetItemInRowCollection(2)->GetCells()->GetItemInCellCollection(0));

	//Save and launch document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}

