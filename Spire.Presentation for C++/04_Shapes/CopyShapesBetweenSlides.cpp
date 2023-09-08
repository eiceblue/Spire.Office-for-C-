#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"CopyShapesBetweenSlides.pptx";
	wstring outputFile = OUTPUTPATH"CopyShapesBetweenSlides.pptx";

	//Load the sample document
	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Define the source slide and target slide
	intrusive_ptr<ISlide> sourceSlide = ppt->GetSlides()->GetItem(0);
	intrusive_ptr<ISlide> targetSlide = ppt->GetSlides()->GetItem(1);

	//Copy the first shape from the source slide to the target slide
	targetSlide->GetShapes()->AddShape(sourceSlide->GetShapes()->GetItem(0));

	//Save the document to file 
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

