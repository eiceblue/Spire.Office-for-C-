#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_1.pptx";
	wstring outputFile = OUTPUTPATH"RemoveTableBorderStyle.pptx";

	//Create a PowerPoint document.
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	for (int l = 0; l < presentation->GetSlides()->GetCount(); l++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(l);
		for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(s);
			intrusive_ptr<ITable> table = Object::Dynamic_cast<ITable> (shape);
			if (table != nullptr)
			{
				for (int i = 0; i < table->GetTableRows()->GetCount(); i++)
				{
					intrusive_ptr<TableRow> row = table->GetTableRows()->GetItem(i);
					for (int j = 0; j < row->GetCount(); j++)
					{
						intrusive_ptr<Cell> cell = row->GetItem(j);
						cell->GetBorderTop()->SetFillType(FillFormatType::None);
						cell->GetBorderBottom()->SetFillType(FillFormatType::None);
						cell->GetBorderLeft()->SetFillType(FillFormatType::None);
						cell->GetBorderRight()->SetFillType(FillFormatType::None);
					}
				}
			}
		}
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
		
}
