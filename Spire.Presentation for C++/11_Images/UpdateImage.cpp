#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"UpdateImage.pptx";
	wstring outputFile = OUTPUTPATH"UpdateImage.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	wstring inputFile1 = DATAPATH"PresentationIcon.png";
	//Append a new image to replace an existing image
	intrusive_ptr<IImageData> image = ppt->GetImages()->Append(new Stream(inputFile1.c_str()));

	//Replace the image which title is "image1" with the new image
	for (int i = 0; i < slide->GetShapes()->GetCount(); i++)
	{
		intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(i);
		if (Object::Dynamic_cast<IEmbedImage>(shape) != nullptr)
		{
			if (wcscmp(shape->GetAlternativeTitle(), L"image1") == 0)
			{
				(Object::Dynamic_cast<IEmbedImage>(shape))->GetPictureFill()->GetPicture()->SetEmbedImage(image);
			}
		}
	}

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

