#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_4.pptx";
	wstring outputFile = OUTPUTPATH"RemoveEncryption.pptx";
	

	//Create a PowerPoint document.
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str(), L"123456");

	//Remove encryption.
	presentation->RemoveEncryption();

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

