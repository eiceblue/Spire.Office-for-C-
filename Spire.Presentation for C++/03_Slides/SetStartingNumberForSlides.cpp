#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChangeSlidePosition.pptx";
	wstring outputFile = OUTPUTPATH"SetStartingNumberForSlides.pptx";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT document from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Set 5 as the starting number
	presentation->SetFirstSlideNumber(5);

	//Save file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
		
}
