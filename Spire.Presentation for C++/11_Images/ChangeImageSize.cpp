#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ExtractImage.pptx";
	wstring outputFile = OUTPUTPATH"ChangeImageSize.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	intrusive_ptr<SlideCollection> slides = ppt->GetSlides();
	float scale = 0.5f;
	for (int i = 0; i < slides->GetCount(); i++)
	{
		intrusive_ptr<ShapeCollection> shapes = slides->GetItem(i)->GetShapes();
		for (int j = 0; j < shapes->GetCount(); j++)
		{
			intrusive_ptr<IShape> shape = shapes->GetItem(j);
			if (Object::Dynamic_cast<IEmbedImage>(shape) != nullptr)
			{
				intrusive_ptr<IEmbedImage> image = Object::Dynamic_cast<IEmbedImage>(shape);
				image->SetWidth(image->GetWidth() * scale);
				image->SetHeight(image->GetHeight() * scale);
			}
		}
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

