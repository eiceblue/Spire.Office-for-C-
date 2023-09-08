#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ArrangeShape.pptx";
	wstring outputFile = OUTPUTPATH"ArrangeShapes.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the specified shape
	intrusive_ptr<IShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0);

	//Bring the shape forward through SetShapeArrange method
	shape->SetShapeArrange(ShapeArrange::BringForward);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

