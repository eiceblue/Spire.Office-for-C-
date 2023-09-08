#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"FindShapeByAltText.pptx";
	wstring outputFile = OUTPUTPATH"FindShapeByAltText.txt";

	wofstream outFile(outputFile, ios::out);

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Find shape in the slide
	intrusive_ptr<IShape> shape = nullptr;

	//Loop through shapes in the slide
	for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape1 = slide->GetShapes()->GetItem(s);
		//Find the shape whose alternative text is altText

		if (wcscmp(shape1->GetAlternativeText(), L"Shape1") == 0)
		{
			shape = shape1;
		}
	}

	outFile << shape->GetName() << endl;
}

