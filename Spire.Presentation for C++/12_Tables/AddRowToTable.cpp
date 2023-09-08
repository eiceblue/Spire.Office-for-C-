#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_1.pptx";
	wstring outputFile = DATAPATH"AddRowToTable.pptx";

	//Create a PowerPoint document.
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the table within the PowerPoint document.
	intrusive_ptr<ITable> table = Object::Dynamic_cast<ITable>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Get the first row.
	intrusive_ptr<TableRow> row = table->GetTableRows()->GetItem(1);

	//Clone the row and add it to the end of table.
	table->GetTableRows()->Append(row);
	int rowCount = table->GetTableRows()->GetCount();

	//Get the last row.
	intrusive_ptr<TableRow> lastRow = table->GetTableRows()->GetItem(rowCount - 1);

	//Set new data of the first cell of last row.
	lastRow->GetItem(0)->GetTextFrame()->SetText(L" The first added cell");

	//Set new data of the second cell of last row.
	lastRow->GetItem(1)->GetTextFrame()->SetText(L" The second added cell");

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
		
}
