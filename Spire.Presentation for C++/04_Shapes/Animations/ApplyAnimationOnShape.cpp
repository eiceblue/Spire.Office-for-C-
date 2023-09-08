#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"ApplyAnimationOnShape.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Set background Image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	slide->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	slide->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Insert a rectangle in the slide and fill the shape
	intrusive_ptr<IAutoShape> shape = slide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(100, 150, 200, 80));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetLightBlue());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->AppendTextFrame(L"Animated Shape");

	//Apply FadedSwivel animation effect to the shape
	shape->GetSlide()->GetTimeline()->GetMainSequence()->AddEffect(shape, AnimationEffectType::FadedSwivel);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}