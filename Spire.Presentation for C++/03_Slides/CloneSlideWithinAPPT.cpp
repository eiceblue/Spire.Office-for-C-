#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"InputTemplate.pptx";
	wstring outputFile = OUTPUTPATH"CloneSlideWithinAPPT.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get a list of slides and choose the first slide to be cloned
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Insert the desired slide to the specified index in the same presentation
	int index = 1;
	ppt->GetSlides()->Insert(index, slide);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	
}
