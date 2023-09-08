#include "pch.h"

using namespace Spire::Presentation;

int main() {

	wstring inputFile = DATAPATH"ExtractImage.pptx";
	wstring outputFile = OUTPUTPATH;
	

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	// Extract Images
	for (int i = 0; i < ppt->GetImages()->GetCount(); i++)
	{
		std::wstring ImageName = outputFile + L"Images_" + to_wstring(i) + L".png";
		intrusive_ptr<Stream> image = ppt->GetImages()->GetItem(i)->GetImage();
		image->Save(ImageName.c_str());
		
	}
}