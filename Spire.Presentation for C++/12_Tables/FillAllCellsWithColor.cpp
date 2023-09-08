#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_1.pptx";
	wstring outputFile = OUTPUTPATH"FillAllCellsWithColor.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Fill the table cells with color.
	intrusive_ptr<ITable> table = nullptr;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		table = Object::Dynamic_cast<ITable> (shape);
		if (table != nullptr)
		{
			for (int t = 0; t < table->GetTableRows()->GetCount(); t++)
			{
				intrusive_ptr<TableRow> row = table->GetTableRows()->GetItem(t);
				for (int n = 0; n < row->GetCount(); n++)
				{
					intrusive_ptr<Cell> cell = row->GetItem(n);
					cell->GetFillFormat()->SetFillType(FillFormatType::Solid);
					cell->GetFillFormat()->GetSolidColor()->SetColor(Color::GetPink());
				}
			}
		}
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
