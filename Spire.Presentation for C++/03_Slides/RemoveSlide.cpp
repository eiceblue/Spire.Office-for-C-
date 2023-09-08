#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"RemoveSlide.pptx";
	wstring outputFile = OUTPUTPATH"RemoveSlide.pptx";


	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Remove slide by index
	presentation->GetSlides()->RemoveAt(0);

	//Remove slide by its reference
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(1);
	presentation->GetSlides()->Remove(slide);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	
}
