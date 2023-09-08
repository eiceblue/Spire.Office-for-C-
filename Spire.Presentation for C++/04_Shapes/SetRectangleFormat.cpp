#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetRectangleFormat.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Add a shape
	intrusive_ptr<RectangleF> rect = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 100, 100, 200, 100);
	intrusive_ptr<IAutoShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rect);

	//Set the fill format of shape
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());

	//Set the fill format of line
	shape->GetLine()->SetFillType(FillFormatType::Solid);
	shape->GetLine()->GetSolidFillColor()->SetColor(Color::GetDimGray());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

