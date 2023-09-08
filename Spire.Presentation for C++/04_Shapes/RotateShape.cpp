#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"RotateShape.pptx";
	wstring outputFile = OUTPUTPATH"RotateShape.pptx";

	//Load a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get the shapes 
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Set the rotation
	shape->SetRotation(60);

	(Object::Dynamic_cast<IAutoShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1)))->SetRotation(120);
	(Object::Dynamic_cast<IAutoShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(2)))->SetRotation(180);
	(Object::Dynamic_cast<IAutoShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(3)))->SetRotation(240);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}

