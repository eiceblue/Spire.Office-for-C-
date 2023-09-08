#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddSection.pptx";
	wstring outputFile = OUTPUTPATH"DeleteSection.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	ppt->GetSectionList()->RemoveAll();

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
