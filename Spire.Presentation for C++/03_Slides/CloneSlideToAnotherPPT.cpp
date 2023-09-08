#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	
	wstring inputFile_1 = DATAPATH"CloneSlideToAnotherPPT-1.pptx";
	wstring inputFile_2 = DATAPATH"CloneSlideToAnotherPPT-2.pptx";
	wstring outputFile = OUTPUTPATH"CloneSlideToAnotherPPT.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile_2.c_str());

	//Load the another document and choose the first slide to be cloned
	intrusive_ptr<Presentation> ppt1 = new Presentation();
	ppt1->LoadFromFile(inputFile_1.c_str());
	intrusive_ptr<ISlide> slide1 = ppt1->GetSlides()->GetItem(0);

	//Insert the slide to the specified index in the source presentation
	int index = 1;
	presentation->GetSlides()->Insert(index, slide1);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	
}
