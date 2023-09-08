#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"CopyParagraph.pptx";
	wstring outputFile = OUTPUTPATH"ConvertPPTToOFD.ofd";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::OFD);
}
