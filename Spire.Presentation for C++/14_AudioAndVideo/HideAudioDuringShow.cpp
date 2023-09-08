#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"audio.pptx";
	wstring outputFile = OUTPUTPATH"HideAudioDuringShow.pptx";
	//Load a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Hide Audio during show
	for (int i = 0; i < slide->GetShapes()->GetCount(); i++)
	{
		intrusive_ptr<IAudio> audio = Object::Dynamic_cast<IAudio>(slide->GetShapes()->GetItem(i));
		if (audio != nullptr)
		{
			audio->SetHideAtShowing(true);
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}