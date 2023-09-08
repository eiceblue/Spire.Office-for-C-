#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Conversion.pptx";
	wstring outputFile = OUTPUTPATH"Image/ToSpecificSizeImage.png";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save the first slide to Image and set the image size to intrusive_ptr<600>400
	intrusive_ptr<Stream> stream = ppt->GetSlides()->GetItem(0)->SaveAsImage(600, 400);
	stream->Save(outputFile.c_str());

}
