#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"RemoveRowsAndColumns.pptx";
	wstring outputFile = OUTPUTPATH"RemoveRowsAndColumns.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the table in PPT document
	intrusive_ptr<ITable> table = nullptr;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		table = Object::Dynamic_cast<ITable> (shape);
		if (table != nullptr)
		{
			//Remove the second column
			table->GetColumnsList()->RemoveAt(1, false);

			//Remove the second row
			table->GetTableRows()->RemoveAt(1, false);
		}
	}
	//Save and launch the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
