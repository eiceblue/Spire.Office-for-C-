#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"InputTemplate.pptx";
	wstring outputFile = OUTPUTPATH"SplitPPT/";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Specify the presentation show type as kiosk
	ppt->SetShowType(SlideShowType::Kiosk);

	intrusive_ptr<SlideCollection> slides = ppt->GetSlides();
	for (int i = 0; i < slides->GetCount(); i++)
	{
		std::wstring res = outputFile + L"SplitPPT-" + to_wstring(i) + L".pptx";
		slides->GetItem(i)->SaveToFile(res.c_str(), FileFormat::Pptx2010);
	}


}
