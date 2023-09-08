#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"sample.dps";
	wstring outputFile = OUTPUTPATH"LoadSaveDPSAndDPT.dps";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str(), FileFormat::Dps);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Dps);
}
