#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChangeSlidePosition.pptx";
	wstring outputFile = OUTPUTPATH"SpecificSlideToPDF.pdf";

	///Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the second slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(1);

	//Save the second slide to PDF
	slide->SaveToFile(outputFile.c_str(), FileFormat::PDF);
}
