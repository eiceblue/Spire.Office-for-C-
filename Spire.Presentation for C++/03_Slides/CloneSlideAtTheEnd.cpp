#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	
	wstring inputFile = DATAPATH"ChangeSlidePosition.pptx";
	wstring outputFile = OUTPUTPATH"CloneSlideAtTheEnd.pptx";

	//Load PPT document from disk
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Append the slide at the end of the document
	presentation->GetSlides()->Append(slide);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
