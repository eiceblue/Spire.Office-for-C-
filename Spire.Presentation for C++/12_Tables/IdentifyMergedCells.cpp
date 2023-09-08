#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"MergedCellInTable.pptx";
	wstring outputFile = OUTPUTPATH"IdentifyMergedCells.txt";

	DeleteFile(outputFile.c_str());
	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);
	wofstream outFile(outputFile.c_str());
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Verify if it is table
		intrusive_ptr<ITable> table = Object::Dynamic_cast<ITable >(shape);
		if (table != nullptr)
		{
			for (int r = 0; r < table->GetTableRows()->GetCount(); r++)
			{
				for (int c = 0; c < table->GetColumnsList()->GetCount(); c++)
				{
					// Get cell
					intrusive_ptr<Cell> currentCell = table->GetTableRows()->GetItem(r)->GetItem(c);
					//Identify if it is merged cell
					if (currentCell->GetRowSpan() > 1 || currentCell->GetColSpan() > 1)
					{
						outFile << "Cell " << to_string(r).c_str() << ":" << to_string(c).c_str() << " is a part of merged cell with RowSpan=" << currentCell->GetRowSpan() << " and ColSpan= " << currentCell->GetColSpan() << " starting from Cell " << currentCell->GetFirstRowIndex() << ":" << currentCell->GetFirstColumnIndex() <<"." << endl;
					}
				}

			}
		}
	}
	outFile.close();
				
}
