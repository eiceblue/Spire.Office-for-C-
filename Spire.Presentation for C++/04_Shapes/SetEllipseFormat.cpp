#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetEllipseFormat.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Add a rectangle
	intrusive_ptr<RectangleF> rect = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 100, 100, 200, 100);
	intrusive_ptr<IAutoShape> shape = slide->GetShapes()->AppendShape(ShapeType::Ellipse, rect);

	//Set the fill format of shape
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());

	//Set the fill format of line
	shape->GetLine()->SetFillType(FillFormatType::Solid);
	shape->GetLine()->GetSolidFillColor()->SetColor(Color::GetDimGray());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

