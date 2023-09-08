#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddSmartArtNode.pptx";
	wstring outputFile = OUTPUTPATH"ChangeSmartArtShapeStyle.pptx";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Get the SmartArt and collect nodes
		intrusive_ptr<ISmartArt> smartArt = Object::Dynamic_cast<ISmartArt>(shape);
		if (smartArt != nullptr)
		{
			//Check SmartArt style
			if (smartArt->GetStyle() == SmartArtStyleType::SimpleFill)
			{
				//Change SmartArt Style
				smartArt->SetStyle(SmartArtStyleType::Cartoon);
			}
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
