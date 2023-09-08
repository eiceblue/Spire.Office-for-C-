#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"GroupShapes.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Set background image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	slide->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	slide->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Create two shapes in the slide
	intrusive_ptr<IShape> rectangle = slide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(250, 180, 200, 40));
	rectangle->GetFill()->SetFillType(FillFormatType::Solid);
	rectangle->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::SkyBlue);
	rectangle->GetLine()->SetWidth(0.1f);
	intrusive_ptr<IShape> ribbon = slide->GetShapes()->AppendShape(ShapeType::Ribbon2, new RectangleF(290, 155, 120, 80));
	ribbon->GetFill()->SetFillType(FillFormatType::Solid);
	ribbon->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::LightPink);
	ribbon->GetLine()->SetWidth(0.1f);

	//Add the two shape objects to an array list
	//intrusive_ptr<ArrayList> list = new ArrayList();
	//list->Add(rectangle);
	//list->Add(ribbon);

	////Group the shapes in the list
	//ppt->GetSlides()->GetItem(0)->SetGroupShapes(list);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

