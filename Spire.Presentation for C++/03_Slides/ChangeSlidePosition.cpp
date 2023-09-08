#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChangeSlidePosition.pptx";
	wstring outputFile = OUTPUTPATH"ChangeSlidePosition.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Move the first slide to the second slide position
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);
	slide->SetSlideNumber(2);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	
}
