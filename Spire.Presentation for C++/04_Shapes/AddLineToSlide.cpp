#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"AddLineToSlide.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Add a line in the slide
	intrusive_ptr<IAutoShape> line = slide->GetShapes()->AppendShape(ShapeType::Line, new RectangleF(50, 100, 300, 0));

	//Set color of the line
	line->GetShapeStyle()->GetLineColor()->SetColor(Color::GetRed());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}

