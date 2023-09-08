#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Macros.ppt";
	wstring outputFile = DATAPATH"RemoveVBAMacros.ppt";

	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Remove macros
	//Note, at present it only can work on macros in PPT file, has not supported for PPTM file yet.
	presentation->DeleteMacros();
	presentation->SaveToFile(outputFile.c_str(), FileFormat::PPT);
}
