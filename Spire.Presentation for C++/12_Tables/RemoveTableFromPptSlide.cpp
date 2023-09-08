#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_1.pptx";
	wstring outputFile = OUTPUTPATH"RemoveTableFromPptSlide.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the tables within the PPT document.
	std::vector<intrusive_ptr<IShape>> shape_tems;

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (Object::Dynamic_cast<ITable>(shape) != nullptr)
		{
			//Add new table to table list.
			shape_tems.push_back(shape);
		}
	}
	//Remove all the tables form the first slide.
	for (auto shape : shape_tems)
	{
		presentation->GetSlides()->GetItem(0)->GetShapes()->Remove(shape);
	}

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
		
}
