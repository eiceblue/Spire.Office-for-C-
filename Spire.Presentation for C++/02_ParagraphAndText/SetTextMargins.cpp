#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetTextMargins.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Append a new shape
	intrusive_ptr<IAutoShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(50, 100, 450, 150));

	//Set margins for text inside shapes
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetLightBlue());
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Justify);
	shape->GetTextFrame()->SetText(L"Using Spire.Presentation, developers will find an easy and effective method to create, read, write, modify, convert and print PowerPoint files on any .Net platform. It's worthwhile for you to try this amazing product.");
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Arial Rounded MT Bold"));
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());

	//Set the margins for the text frame
	shape->GetTextFrame()->SetMarginTop(10);
	shape->GetTextFrame()->SetMarginBottom(35);
	shape->GetTextFrame()->SetMarginLeft(15);
	shape->GetTextFrame()->SetMarginRight(30);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

