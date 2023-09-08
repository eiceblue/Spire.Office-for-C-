#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile =DATAPATH"video.pptx";
	wstring outputFile = OUTPUTPATH"ExtractVideo/";

	//Load a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());
	intrusive_ptr<SlideCollection> slides = presentation->GetSlides();

	int index = 0;
	for (int i = 0; i < slides->GetCount(); i++)
	{
		intrusive_ptr<ShapeCollection> shapes = slides->GetItem(i)->GetShapes();
		for (int j = 0; j < shapes->GetCount(); j++)
		{
			intrusive_ptr<IVideo> audio = Object::Dynamic_cast<IVideo>(shapes->GetItem(j));
			if (audio != nullptr)
			{
				audio->GetEmbeddedVideoData()->SaveToFile((outputFile + L"ExtractVideo_" + to_wstring(index) + L".avi").c_str());
				index++;
			}
		}
	}
}