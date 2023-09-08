#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"GetShapeGroupAltText.pptx";
	wstring outputFile = OUTPUTPATH"GetShapeAltText.txt";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load document from disk
	presentation->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);

	//Loop through slides and shapes
	for (int s = 0; s < presentation->GetSlides()->GetCount(); s++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(s);
		for (int i = 0; i < slide->GetShapes()->GetCount(); i++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(i);
			if (Object::Dynamic_cast<GroupShape>(shape) != nullptr)
			{
				//Find the shape group
				intrusive_ptr<GroupShape> groupShape = Object::Dynamic_cast<GroupShape>(shape);
				for (int p = 0; p < groupShape->GetShapes()->GetCount(); p++)
				{
					intrusive_ptr<IShape> gShape = groupShape->GetShapes()->GetItem(p);
					//Append the alternative text in builder
					outFile << gShape->GetAlternativeText() << endl;
				}
			}
		}
	}

	//Write the content in txt file
	outFile.close();
}

