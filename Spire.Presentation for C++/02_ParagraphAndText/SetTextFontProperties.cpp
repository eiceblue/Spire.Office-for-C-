#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetTextFontProperties.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Add a new shape to the PPT document
	intrusive_ptr<RectangleF> rec = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 250, 80, 500, 150);
	intrusive_ptr<IAutoShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec);

	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);

	//Add text to the shape
	shape->AppendTextFrame(L"Welcome to use Spire.Presentation");

	intrusive_ptr<TextRange> textRange = shape->GetTextFrame()->GetTextRange();
	//Set the font
	textRange->SetLatinFont(new TextFont(L"Times New Roman"));
	//Set bold property of the font
	textRange->SetIsBold(TriState::True);

	//Set italic property of the font
	textRange->SetIsItalic(TriState::True);

	//Set underline property of the font
	textRange->SetTextUnderlineType(TextUnderlineType::Single);

	//Set the height of the font
	textRange->SetFontHeight(50);

	//Set the color of the font
	textRange->GetFill()->SetFillType(FillFormatType::Solid);
	textRange->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

