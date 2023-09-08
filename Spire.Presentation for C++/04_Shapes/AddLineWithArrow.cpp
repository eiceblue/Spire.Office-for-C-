#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"AddLineWithArrow.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Add a line to the slides and set its color to red
	intrusive_ptr<IAutoShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Line, new RectangleF(150, 100, 100, 100));
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetRed());
	//Set the line end type as StealthArrow
	shape->GetLine()->SetLineEndType(LineEndType::StealthArrow);

	//Add a line to the slides and use default color
	shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Line, new RectangleF(300, 150, 100, 100));
	shape->SetRotation(-45);
	//Set the line end type as TriangleArrowHead
	shape->GetLine()->SetLineEndType(LineEndType::TriangleArrowHead);

	//Add a line to the slides and set its color to Green
	shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Line, new RectangleF(450, 100, 100, 100));
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetGreen());
	shape->SetRotation(90);
	//Set the line begin type as TriangleArrowHead
	shape->GetLine()->SetLineBeginType(LineEndType::StealthArrow);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

