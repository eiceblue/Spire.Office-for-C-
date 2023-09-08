
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_8.pptx";
	wstring outputFile = OUTPUTPATH"SetPlayModeForVideo.pptx";
	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT document from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Find the video by looping through all the slides and set its play mode as auto.
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(i);
		intrusive_ptr<ShapeCollection> shapes = slide->GetShapes();
		for (int j = 0; j < shapes->GetCount(); j++)
		{
			intrusive_ptr<IShape> shape = shapes->GetItem(j);
			//If shape is IVideo
			//Replace the video
			intrusive_ptr<IVideo> video = Object::Dynamic_cast<IVideo>(shape);
			if (video != nullptr)
			{
				video->SetPlayMode(VideoPlayMode::Auto);
			}
		}
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
