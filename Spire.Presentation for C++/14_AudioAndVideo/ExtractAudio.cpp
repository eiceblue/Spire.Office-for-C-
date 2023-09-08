#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"audio.pptx";
	wstring outputFile = OUTPUTPATH"ExtractAudio.wav";
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	intrusive_ptr<ShapeCollection> shapes = presentation->GetSlides()->GetItem(0)->GetShapes();
	int index = 1;
	for (int i = 0; i < shapes->GetCount(); i++)
	{
		intrusive_ptr<IAudio> audio = Object::Dynamic_cast<IAudio>(shapes->GetItem(i));
		if (audio != nullptr)
		{
			audio->GetData()->SaveToFile(outputFile.c_str());
			index++;
		}
	}
}