#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_1.pptx";
	wstring outputFile = OUTPUTPATH"TraverseThroughCells.txt";

	//Create a PowerPonit document.
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	wofstream content(outputFile);
	content << "The data in cells of this PowerPoint file is: " << endl;

	//Get the table.
	intrusive_ptr<ITable> table = nullptr;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		table = Object::Dynamic_cast<ITable >(shape);
		if (table != nullptr)
		{
			//Traverse through the cells of table.
			for (int r = 0; r < table->GetTableRows()->GetCount(); r++)
			{
				intrusive_ptr<TableRow> row = table->GetTableRows()->GetItem(r);
				for (int c = 0; c < row->GetCount(); c++)
				{
					intrusive_ptr<Cell> cell = row->GetItem(c);
					content << cell->GetTextFrame()->GetText() << endl;
				}
				content << endl;
			}
		}
	}
	//Save to file.
	content.close();
}
