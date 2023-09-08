#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"MergeTableCell.pptx";
	wstring outputFile = OUTPUTPATH"MergeTableCell.pptx";

	//Create a PPT document and load file
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	intrusive_ptr<ITable> table = nullptr;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		table = Object::Dynamic_cast<ITable> (shape);
		if (table != nullptr)
		{
			//Merge the second row and third row of the first column
			table->MergeCells(table->GetItem(0, 1), table->GetItem(0, 2), false);

			table->MergeCells(table->GetItem(3, 4), table->GetItem(4, 4), true);
		}
	}
	//Save and launch the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
			
}
