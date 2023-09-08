#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"AddShapes.pptx";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Set background Image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, presentation->GetSlideSize()->GetSize()->GetWidth(), presentation->GetSlideSize()->GetSize()->GetHeight());
	presentation->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Append new shape - Triangle and set style
	intrusive_ptr<IAutoShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Triangle, new RectangleF(115, 130, 100, 100));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetLightGreen());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());

	//Append new shape - Ellipse
	shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Ellipse, new RectangleF(290, 130, 150, 100));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetLightSkyBlue());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());

	//Append new shape - Heart
	shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Heart, new RectangleF(470, 130, 130, 100));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetRed());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetLightGray());


	//Append new shape - FivePointedStar
	shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::FivePointedStar, new RectangleF(90, 270, 150, 150));
	shape->GetFill()->SetFillType(FillFormatType::Gradient);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());

	//Append new shape - Rectangle
	shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(320, 290, 100, 120));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetPink());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetLightGray());

	//Append new shape - BentUpArrow
	shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::BentUpArrow, new RectangleF(470, 300, 150, 100));

	//Set the color of shape
	shape->GetFill()->SetFillType(FillFormatType::Gradient);
	shape->GetFill()->GetGradient()->GetGradientStops()->Append(1.0f, KnownColors::Olive);
	shape->GetFill()->GetGradient()->GetGradientStops()->Append(0, KnownColors::PowderBlue);
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}

