#include "pch.h"

using namespace Spire::Presentation;

int main()
{

	wstring inputFile = DATAPATH"FindShapeByAltText.pptx";
	wstring outputFile = OUTPUTPATH"RemoveShape.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load doucment from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Loop through slides
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(i);
		//Loop through shapes
		for (int j = 0; j < slide->GetShapes()->GetCount(); j++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(j);
			//Find the shapes whose alternative text contain "Shape"
			wstring temp = shape->GetAlternativeText();
			std::string::size_type pos = temp.find(L"Shape");
			if (pos != string::npos)
			{
				slide->GetShapes()->Remove(shape);
				j--;
			}
		}
	}

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

