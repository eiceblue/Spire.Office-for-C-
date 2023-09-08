#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChangeSlidePosition.pptx";
	wstring outputFile = OUTPUTPATH"IndividualSlideToHtml.html";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Save the first slide to HTML 
	slide->SaveToFile(outputFile.c_str(), FileFormat::Html);
}
