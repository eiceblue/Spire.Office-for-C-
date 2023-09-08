#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SetRowHeightColumnWidth.pptx";
	wstring outputFile = OUTPUTPATH"SetRowHeightColumnWidth.pptx";

	//Creat a ppt document and load file
	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get the table
	intrusive_ptr<ITable> table = nullptr;
	for (int s = 0; s < ppt->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		table = Object::Dynamic_cast<ITable > (shape);
		if (table != nullptr)
		{
			//Set the height for the rows
			table->GetTableRows()->GetItem(0)->SetHeight(100);
			table->GetTableRows()->GetItem(1)->SetHeight(80);
			table->GetTableRows()->GetItem(2)->SetHeight(60);
			table->GetTableRows()->GetItem(3)->SetHeight(40);
			table->GetTableRows()->GetItem(4)->SetHeight(20);

			//Set the column width
			table->GetColumnsList()->GetItem(0)->SetWidth(60);
			table->GetColumnsList()->GetItem(1)->SetWidth(80);
			table->GetColumnsList()->GetItem(2)->SetWidth(120);
			table->GetColumnsList()->GetItem(3)->SetWidth(140);
			table->GetColumnsList()->GetItem(4)->SetWidth(160);
		}
	}
	//Save the file
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
		
}
