
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{

	wstring inputFile_1 = DATAPATH"video.pptx";
	wstring inputFile_2 = DATAPATH"repleaceVido.mp4";
	wstring outputFile = OUTPUTPATH"ReplaceVideo.pptx";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT document from disk.
	presentation->LoadFromFile(inputFile_1.c_str());

	intrusive_ptr<VideoCollection> videos = presentation->GetVideos();

	//Traverse all the slides of PPT file
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(i);
		intrusive_ptr<ShapeCollection> shapes = slide->GetShapes();
		//Traverse all the shapes of slides
		for (int j = 0; j < shapes->GetCount(); j++)
		{
			intrusive_ptr<IShape> shape = shapes->GetItem(j);
			//If shape is IVideo
			//Replace the video
			intrusive_ptr<IVideo> video = Object::Dynamic_cast<IVideo>(shape);
			if (video)
			{
				//Load the video document from disk.
				intrusive_ptr<Stream> videoStream = new Stream(inputFile_2.c_str());
				intrusive_ptr<VideoData> videoData = videos->Append(videoStream);
				video->SetEmbeddedVideoData(videoData);
			}
		}
	}
	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
