#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"OneSlideToSVG.pptx";
	wstring outputFile = OUTPUTPATH"SVG/OneSlideToSVG/OneSlideToSVG.svg";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Convert the second slide to SVG
	intrusive_ptr<Stream> svg = ppt->GetSlides()->GetItem(1)->SaveToSVG();
	svg->Save(outputFile.c_str());
	svg->Close();

}
