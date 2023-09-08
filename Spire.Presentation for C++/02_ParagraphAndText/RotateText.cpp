#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Az1.pptx";
	wstring outputFile = OUTPUTPATH"RotateText.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);
	//Get a shape 
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	shape->GetTextFrame()->SetVerticalTextType(VerticalTextType::Vertical270);

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}

