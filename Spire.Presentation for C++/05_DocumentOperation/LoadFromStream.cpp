#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"InputTemplate.pptx";
	wstring outputFile = OUTPUTPATH"LoadFromStream.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	intrusive_ptr<Stream> stream = new Stream(inputFile.c_str());
	ppt->LoadFromStream(stream, FileFormat::Pptx2013);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
