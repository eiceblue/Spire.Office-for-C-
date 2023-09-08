#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"HelloWorld.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Add a new shape to the PPT document
	intrusive_ptr<RectangleF> rec = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 250, 80, 500, 150);
	intrusive_ptr<IAutoShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec);

	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);

	//Add text to the shape
	shape->AppendTextFrame(L"Hello World!");

	//Set the font and fill style of the text
	intrusive_ptr<TextRange> textRange = shape->GetTextFrame()->GetTextRange();
	textRange->GetFill()->SetFillType(FillFormatType::Solid);
	textRange->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());
	textRange->SetFontHeight(66);
	textRange->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

}

