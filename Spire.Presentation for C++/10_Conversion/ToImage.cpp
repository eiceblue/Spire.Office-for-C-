#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ToImage.pptx";
	wstring outputFile = OUTPUTPATH"Image/ToImage/";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save PPT document to images
	intrusive_ptr<SlideCollection> slides = ppt->GetSlides();
	for (int i = 0; i < slides->GetCount(); i++)
	{
		intrusive_ptr<ISlide> slide = slides->GetItem(i);
		intrusive_ptr<Stream> image = slide->SaveAsImage();
		image->Save((outputFile + L"ToImage_img_" + to_wstring(i) + L".png").c_str());
	}
}
