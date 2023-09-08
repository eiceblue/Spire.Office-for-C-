#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"OpenEncryptedPPT.pptx";
	wstring outputFile = OUTPUTPATH"OpenEncryptedPPT.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT with password
	presentation->LoadFromFile(inputFile.c_str(), FileFormat::Pptx2010, L"123456");

	//Save as a new PPT with original password
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}

