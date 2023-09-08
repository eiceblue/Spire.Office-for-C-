#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"FillShapeWithSolidColor.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Add a rectangle
	intrusive_ptr<RectangleF> rect = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 50, 100, 100, 100);
	intrusive_ptr<IAutoShape> shape = slide->GetShapes()->AppendShape(ShapeType::Rectangle, rect);

	//Fill shape with solid color
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetYellow());

	//Set the fill format of line
	shape->GetLine()->SetFillType(FillFormatType::Solid);
	shape->GetLine()->GetSolidFillColor()->SetColor(Color::GetGray());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

