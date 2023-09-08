#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Images.pptx";
	wstring outputFile = OUTPUTPATH;
	

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the pictures on the second slide and save them to image file
	int i = 0;
	intrusive_ptr<ShapeCollection> shapes = ppt->GetSlides()->GetItem(1)->GetShapes();
	//Traverse all shapes in the second slide
	for (int j = 0; j < shapes->GetCount(); j++)
	{
		intrusive_ptr<IShape> s = shapes->GetItem(j);
		//It is the SlidePicture object
		if (Object::Dynamic_cast<IEmbedImage>(s) != nullptr)
		{
			//Save to image
			intrusive_ptr<IEmbedImage> ps = Object::Dynamic_cast<IEmbedImage>(s);
			ps->GetPictureFill()->GetPicture()->GetEmbedImage()->GetImage()->Save((outputFile + L"SlidePic_" + to_wstring(i) + L".png").c_str());
			i++;
		}
		//It is the PictureShape object
		if (Object::Dynamic_cast<PictureShape>(s) != nullptr)
		{
			//Save to image
			intrusive_ptr<PictureShape> ps = Object::Dynamic_cast<PictureShape>(s);
			ps->GetEmbedImage()->GetImage()->Save((outputFile + L"ShapePic_" + to_wstring(i) + L".png").c_str());
			i++;
		}
	}
}

