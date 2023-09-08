#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateSlide.pptx";
	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Add new slide
	presentation->GetSlides()->Append();

	//Set the background image
	for (int i = 0; i < 2; i++)
	{
		wstring ImageFile = DATAPATH"bg.png";
		intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, presentation->GetSlideSize()->GetSize()->GetWidth(), presentation->GetSlideSize()->GetSize()->GetHeight());
		presentation->GetSlides()->GetItem(i)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
		presentation->GetSlides()->GetItem(i)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());
	}

	//Add title
	intrusive_ptr<RectangleF> rec_title = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 200, 70, 400, 50);
	intrusive_ptr<IAutoShape> shape_title = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec_title);
	shape_title->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape_title->GetFill()->SetFillType(FillFormatType::None);
	intrusive_ptr<TextParagraph> para_title = new TextParagraph();
	para_title->SetText(L"E-iceblue");
	para_title->SetAlignment(TextAlignmentType::Center);
	para_title->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Myriad Pro Light"));
	para_title->GetTextRanges()->GetItem(0)->SetFontHeight(36);
	para_title->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	para_title->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());
	shape_title->GetTextFrame()->GetParagraphs()->Append(para_title);

	//Append new shape
	intrusive_ptr<IAutoShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(50, 150, 600, 280));
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetLine()->SetFillType(FillFormatType::None);
	//Add text to shape
	shape->AppendTextFrame(L"Welcome to use Spire.Presentation for .NET.");

	//Add new paragraph
	intrusive_ptr<TextParagraph> pare = new TextParagraph();
	pare->SetText(L"");
	shape->GetTextFrame()->GetParagraphs()->Append(pare);

	//Add new paragraph
	pare = new TextParagraph();
	pare->SetText(L"Spire.Presentation for .NET is a professional PowerPoint compatible component that enables developers to create, read, write, modify, convert and Print PowerPoint documents from any .NET(C#, VB.NET, ASP.NET) platform. As an independent PowerPoint .NET component, Spire.Presentation for .NET doesn't need Microsoft PowerPoint installed on the machine.");
	shape->GetTextFrame()->GetParagraphs()->Append(pare);

	//Set the Font
	for (int t = 0; t < shape->GetTextFrame()->GetParagraphs()->GetCount(); t++)
	{
		intrusive_ptr<TextParagraph> para = shape->GetTextFrame()->GetParagraphs()->GetItem(t);
		para->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Myriad Pro"));
		para->GetTextRanges()->GetItem(0)->SetFontHeight(24);
		para->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
		para->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());
		para->SetAlignment(TextAlignmentType::Left);
	}

	//Append new shape - SixPointedStar
	shape = presentation->GetSlides()->GetItem(1)->GetShapes()->AppendShape(ShapeType::SixPointedStar, new RectangleF(100, 100, 100, 100));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetOrange());
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());

	//Append new shape
	shape = presentation->GetSlides()->GetItem(1)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(50, 250, 600, 50));
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);

	//Add text to shape
	shape->AppendTextFrame(L"This is newly added Slide.");

	//Set the Font
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Myriad Pro"));
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetFontHeight(24);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Left);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetIndent(35);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
			
}
