#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ShapeTemplate.pptx";
	wstring outputFile = OUTPUTPATH"ResetShapeSizeAndPosition.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Define the original slide size
	float currentHeight = ppt->GetSlideSize()->GetSize()->GetHeight();
	float currentWidth = ppt->GetSlideSize()->GetSize()->GetWidth();

	//Change the slide size as A3
	ppt->GetSlideSize()->SetType(SlideSizeType::A3);

	//Define the new slide size
	float newHeight = ppt->GetSlideSize()->GetSize()->GetHeight();
	float newWidth = ppt->GetSlideSize()->GetSize()->GetWidth();

	//Define the ratio from the old and new slide size
	float ratioHeight = newHeight / currentHeight;
	float ratioWidth = newWidth / currentWidth;

	//Reset the size and position of the shape on the slide
	for (int l = 0; l < ppt->GetSlides()->GetCount(); l++)
	{
		intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(l);
		for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(s);
			shape->SetHeight(shape->GetHeight() * ratioHeight);
			shape->SetWidth(shape->GetWidth() * ratioWidth);

			shape->SetLeft(shape->GetLeft() * ratioHeight);
			shape->SetTop(shape->GetTop() * ratioWidth);
		}
	}
	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}

