#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Animations.pptx";
	wstring outputFile = OUTPUTPATH"Animations.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Add title
	intrusive_ptr<RectangleF> rec_title = new RectangleF(50, 200, 200, 50);
	intrusive_ptr<IAutoShape> shape_title = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec_title);
	shape_title->GetShapeStyle()->GetLineColor()->SetColor(Color::GetTransparent());

	shape_title->GetFill()->SetFillType(FillFormatType::None);
	intrusive_ptr<TextParagraph> para_title = new TextParagraph();
	para_title->SetText(L"Animations:");
	para_title->SetAlignment(TextAlignmentType::Center);
	para_title->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Myriad Pro Light"));
	para_title->GetTextRanges()->GetItem(0)->SetFontHeight(32);
	para_title->GetTextRanges()->GetItem(0)->SetIsBold(TriState::True);
	para_title->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	para_title->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::FromArgb(68, 68, 68));
	shape_title->GetTextFrame()->GetParagraphs()->Append(para_title);

	//Set the animation of slide to Circle
	presentation->GetSlides()->GetItem(0)->GetSlideShowTransition()->SetType(TransitionType::Circle);

	//Append new shape - Triangle
	intrusive_ptr<IAutoShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Triangle, new RectangleF(100, 280, 80, 80));

	//Set the color of shape
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());

	//Set the animation of shape
	shape->GetSlide()->GetTimeline()->GetMainSequence()->AddEffect(shape, AnimationEffectType::Path4PointStar);

	//Append new shape - Rectangle and set animation
	shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(210, 280, 150, 80));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->AppendTextFrame(L"Animated Shape");
	shape->GetSlide()->GetTimeline()->GetMainSequence()->AddEffect(shape, AnimationEffectType::FadedSwivel);

	//Append new shape - Cloud and set the animation
	shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Cloud, new RectangleF(390, 280, 80, 80));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetWhite());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetCadetBlue());
	shape->GetSlide()->GetTimeline()->GetMainSequence()->AddEffect(shape, AnimationEffectType::FadedZoom);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}

