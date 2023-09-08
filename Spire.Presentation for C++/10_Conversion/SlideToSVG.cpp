#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"PPTSample_N.pptx";
	wstring outputFile = OUTPUTPATH"SVG/SlideToSVG/SlideToSVG.svg";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Convert the second slide to SVG
	intrusive_ptr<Stream> svg = ppt->GetSlides()->GetItem(0)->SaveToSVG();
	svg->Save(outputFile.c_str());
	svg->Close();
}
