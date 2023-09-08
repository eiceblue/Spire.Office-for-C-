#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChangeSlideLayout.pptx";
	wstring outputFile = OUTPUTPATH"ChangeSlideLayout.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Change the layout of slide
	presentation->GetSlides()->GetItem(1)->SetLayout(presentation->GetMasters()->GetItem(0)->GetLayouts()->GetItem(4));

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

		
}
