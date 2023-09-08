#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetRadiusOfRoundedRectangle.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Insert a rounded rectangle and set its radious
	presentation->GetSlides()->GetItem(0)->GetShapes()->InsertRoundRectangle(0, 160, 180, 100, 200, 10);

	//Append a rounded rectangle and set its radius
	intrusive_ptr<IAutoShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendRoundRectangle(380, 180, 100, 200, 100);
	//Set the color and fill style of shape
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetSeaGreen());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());

	//Rotate the shape to 90 degree
	shape->SetRotation(90);

	//Save the document to Pptx file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

