#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"PageSetup.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();


	//Set the size of slides
	ppt->GetSlideSize()->SetSize(new SizeF(600, 600));
	ppt->GetSlideSize()->SetOrientation(SlideOrienation::Portrait);
	ppt->GetSlideSize()->SetType(SlideSizeType::Custom);

	//Set background image
	std::wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Append new shape
	intrusive_ptr<RectangleF> rec = new RectangleF(ppt->GetSlideSize()->GetSize()->GetWidth() / 2 - 200, 150, 400, 200);
	intrusive_ptr<IAutoShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec);
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);

	//Add text to shape
	shape->AppendTextFrame(L"The sample demonstrates how to set slide size.");

	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Myriad Pro"));
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetFontHeight(24);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::FromArgb(36, 64, 97));

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

}
