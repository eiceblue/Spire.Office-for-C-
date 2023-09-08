#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddWatermark.pptx";
	wstring outputFile = OUTPUTPATH"Watermark.pptx";

	//Create a PPT document and load file
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Define a rectangle range
	intrusive_ptr<RectangleF> rect = new RectangleF((presentation->GetSlideSize()->GetSize()->GetWidth() - 337) / 2, (presentation->GetSlideSize()->GetSize()->GetHeight() - 111) / 2, 337, 111);

	//Add a rectangle shape with a defined range
	intrusive_ptr<IAutoShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rect);

	//Set the style of the shape
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->SetRotation(-45);
	shape->GetLocking()->SetSelectionProtection(true);
	shape->GetLine()->SetFillType(FillFormatType::None);

	//Add text to the shape
	shape->GetTextFrame()->SetText(L"E-iceblue");
	intrusive_ptr<TextRange> textRange = shape->GetTextFrame()->GetTextRange();
	//Set the style of the text range
	textRange->GetFill()->SetFillType(FillFormatType::Solid);
	textRange->GetFill()->GetSolidColor()->SetColor(Color::FromArgb(120, Color::GetHotPink()));
	textRange->SetFontHeight(50);

	//Save the document and launch
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
		
}


