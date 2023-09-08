#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"InputTemplate.pptx";
	wstring outputFile = OUTPUTPATH"LoopPresentPPT.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Set the Boolean value of ShowLoop as true
	ppt->SetShowLoop(true);

	//Set the PowerPoint document to show animation and narration
	ppt->SetShowAnimation(true);
	ppt->SetShowNarration(true);
	//Use slide transition timings to advance slide
	ppt->SetUseTimings(true);

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
