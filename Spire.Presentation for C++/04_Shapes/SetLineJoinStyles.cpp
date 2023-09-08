#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetLineJoinStyles.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Add three shapes
	intrusive_ptr<IAutoShape> shape1 = slide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(50, 150, 150, 50));
	intrusive_ptr<IAutoShape> shape2 = slide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(250, 150, 150, 50));
	intrusive_ptr<IAutoShape> shape3 = slide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(450, 150, 150, 50));

	//Fill shapes
	shape1->GetFill()->SetFillType(FillFormatType::Solid);
	shape1->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());
	shape2->GetFill()->SetFillType(FillFormatType::Solid);
	shape2->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());
	shape3->GetFill()->SetFillType(FillFormatType::Solid);
	shape3->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());

	//Fill lines of shapes
	shape1->GetLine()->SetFillType(FillFormatType::Solid);
	shape1->GetLine()->GetSolidFillColor()->SetColor(Color::GetDarkGray());
	shape2->GetLine()->SetFillType(FillFormatType::Solid);
	shape2->GetLine()->GetSolidFillColor()->SetColor(Color::GetDarkGray());
	shape3->GetLine()->SetFillType(FillFormatType::Solid);
	shape3->GetLine()->GetSolidFillColor()->SetColor(Color::GetDarkGray());

	//Set the line width
	shape1->GetLine()->SetWidth(10);
	shape2->GetLine()->SetWidth(10);
	shape3->GetLine()->SetWidth(10);

	//Set the join styles of lines
	shape1->GetLine()->SetJoinStyle(LineJoinType::Bevel);
	shape2->GetLine()->SetJoinStyle(LineJoinType::Miter);
	shape3->GetLine()->SetJoinStyle(LineJoinType::Round);

	//Add text in shapes
	shape1->GetTextFrame()->SetText(L"Bevel Join Style");
	shape2->GetTextFrame()->SetText(L"Miter Join Style");
	shape3->GetTextFrame()->SetText(L"Round Join Style");

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

