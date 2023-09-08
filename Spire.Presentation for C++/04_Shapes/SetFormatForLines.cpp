#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetFormatForLines.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Add a rectangle shape to the slide
	intrusive_ptr<IAutoShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(100, 150, 200, 100));
	//Set the fill color of the rectangle shape
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetWhite());
	//Apply some formatting on the line of the rectangle
	shape->GetLine()->SetStyle(TextLineStyle::ThickThin);
	shape->GetLine()->SetWidth(5);
	shape->GetLine()->SetDashStyle(LineDashStyleType::Dash);
	//Set the color of the line of the rectangle
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetSkyBlue());

	//Add a ellipse shape to the slide
	shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Ellipse, new RectangleF(400, 150, 200, 100));
	//Set the fill color of the ellipse shape
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetWhite());
	//Apply some formatting on the line of the ellipse
	shape->GetLine()->SetStyle(TextLineStyle::ThickBetweenThin);
	shape->GetLine()->SetWidth(5);
	shape->GetLine()->SetDashStyle(LineDashStyleType::DashDot);
	//Set the color of the line of the ellipse
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetOrangeRed());

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}

