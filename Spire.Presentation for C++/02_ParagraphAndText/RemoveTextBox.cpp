#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"TextBoxTemplate.pptx";
	wstring outputFile = OUTPUTPATH"RemoveTextBox.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);
	//Traverse all the shapes in slide
	for (int i = 0; i < slide->GetShapes()->GetCount();)
	{
		//Remove all shapes
		intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(slide->GetShapes()->GetItem(i));
		slide->GetShapes()->Remove(shape);
	}

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

