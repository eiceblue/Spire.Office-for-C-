#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"RemoveImages.pptx";
	wstring outputFile = OUTPUTPATH"RemoveImages.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	for (int i = slide->GetShapes()->GetCount() - 1; i >= 0; i--)
	{
		//It is the SlidePicture object
		if (Object::Dynamic_cast<IEmbedImage>(slide->GetShapes()->GetItem(i)) != nullptr)
		{
			slide->GetShapes()->RemoveAt(i);
		}
	}
	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

