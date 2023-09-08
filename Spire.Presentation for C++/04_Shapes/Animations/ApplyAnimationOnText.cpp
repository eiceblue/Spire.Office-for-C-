#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"ApplyAnimationOnText.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Set background image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	slide->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	slide->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Add a shape to the slide
	intrusive_ptr<IAutoShape> shape = slide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(250, 150, 200, 100));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetLightBlue());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->AppendTextFrame(L"This demo shows how to apply animation on text in PPT document.");

	//Apply animation to the text in shape
	intrusive_ptr<AnimationEffect> animation = shape->GetSlide()->GetTimeline()->GetMainSequence()->AddEffect(shape, AnimationEffectType::Float);
	animation->SetStartEndParagraphs(0, 0);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

