#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"AddDifferentLayoutSlides.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Remove the default slide
	presentation->GetSlides()->RemoveAt(0);

	//Loop through slide layouts
	for (int i = (int)SlideLayoutType::Title; i <= (int)SlideLayoutType::PictureAndCaption; i++)
	{
		SlideLayoutType type = static_cast<SlideLayoutType>(i);
		presentation->GetSlides()->Append(type);
	}

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
				
	
}
