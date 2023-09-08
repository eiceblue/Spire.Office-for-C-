#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"RemoveAllDigitalSignatures.pptx";
	wstring outputFile = OUTPUTPATH"RemoveAllDigitalSignatures.pptx";

	//Create a PowerPoint document.
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Remove all digital signatures
	if (ppt->GetIsDigitallySigned())
	{
		ppt->RemoveAllDigitalSignatures();
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

}

