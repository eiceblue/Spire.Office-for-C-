#include "pch.h"
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateSlideMasterAndApply.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	ppt->GetSlideSize()->SetType(SlideSizeType::Screen16x9);

	//Add slides
	for (int i = 0; i < 4; i++)
	{
		ppt->GetSlides()->Append();
	}

	//Get the first default slide master
	intrusive_ptr<IMasterSlide> first_master = ppt->GetMasters()->GetItem(0);

	//Append another slide master
	ppt->GetMasters()->AppendSlide(first_master);
	intrusive_ptr<IMasterSlide> second_master = ppt->GetMasters()->GetItem(1);

	//Set different background image for the two slide masters
	wstring pic1 = DATAPATH"bg.png";
	wstring pic2 = DATAPATH"Setbackground.png";
	//The first slide master
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	first_master->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Picture);
	intrusive_ptr<IEmbedImage> image1 = first_master->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, pic1.c_str(), rect);
	first_master->GetSlideBackground()->GetFill()->GetPictureFill()->GetPicture()->SetEmbedImage(image1->GetPictureFill()->GetPicture()->GetEmbedImage());
	//The second slide master
	second_master->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Picture);
	intrusive_ptr<IEmbedImage> image2 = second_master->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, pic2.c_str(), rect);
	second_master->GetSlideBackground()->GetFill()->GetPictureFill()->GetPicture()->SetEmbedImage(image2->GetPictureFill()->GetPicture()->GetEmbedImage());

	//Apply the first master with layout to the first slide
	ppt->GetSlides()->GetItem(0)->SetLayout(first_master->GetLayouts()->GetItem(1));

	//Apply the second master with layout to other slides
	for (int i = 1; i < ppt->GetSlides()->GetCount(); i++)
	{
		ppt->GetSlides()->GetItem(i)->SetLayout(second_master->GetLayouts()->GetItem(8));
	}

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
				
}
