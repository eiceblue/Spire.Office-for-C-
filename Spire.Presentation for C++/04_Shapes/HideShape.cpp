#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"FindShapeByAltText.pptx";
	wstring outputFile = OUTPUTPATH"HideShape.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Loop through slides
	for (int l = 0; l < presentation->GetSlides()->GetCount(); l++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(l);
		//Loop through shapes in the slide
		for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(s);
			//Find the shape whose alternative text is Shape1
			if (wcscmp(shape->GetAlternativeText(), L"Shape1") == 0)
			{
				//Hide the shape
				shape->SetIsHidden(true);
			}
		}
	}
	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

