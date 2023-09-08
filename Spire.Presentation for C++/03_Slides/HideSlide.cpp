#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"HideSlide.pptx";
	wstring outputFile = OUTPUTPATH"HideSlide.pptx";

	//Create a PPT document and load PPT file from disk
	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Hide the second slide
	ppt->GetSlides()->GetItem(1)->SetHidden(true);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	
}
