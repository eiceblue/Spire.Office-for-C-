#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"InputTemplate.pptx";
	wstring outputFile = OUTPUTPATH"SlideTitle.pptx";


	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);
	//Get the title of the first slide
	std::wstring slideTitle = slide->GetTitle();
	//Set the title of the second slide
	ppt->GetSlides()->GetItem(1)->SetTitle(L"Second Slide");
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
				
	
}
