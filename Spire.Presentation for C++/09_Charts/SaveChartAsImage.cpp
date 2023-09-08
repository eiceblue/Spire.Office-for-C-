
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SaveChartAsImage.pptx";
	wstring outputFile = OUTPUTPATH"Image/SaveChartAsImage.png";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save chart as image in .png format
	intrusive_ptr<Stream> image = ppt->GetSlides()->GetItem(0)->GetShapes()->SaveAsImage(0);
std:ofstream output(outputFile.c_str(), ios::binary);
	image->Save(output);
	output.close();
	image->Close();
}
