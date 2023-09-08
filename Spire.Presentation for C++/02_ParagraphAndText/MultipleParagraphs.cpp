#include "pch.h"

using namespace Spire::Presentation;

int main()
{

	wstring inputFile = DATAPATH"Template_Az.pptx";
	wstring outputFile = OUTPUTPATH"MultipleParagraphs.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Access the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	// Add an AutoShape of rectangle type
	intrusive_ptr<RectangleF> rec = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 250, 150, 500, 150);
	intrusive_ptr<IAutoShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec);

	// Access TextFrame of the AutoShape
	intrusive_ptr<ITextFrameProperties> tf = shape->GetTextFrame();

	// Create Paragraphs and TextRanges with different text formats
	intrusive_ptr<TextParagraph> para0 = tf->GetParagraphs()->GetItem(0);
	intrusive_ptr<TextRange> textRange1 = new TextRange();
	intrusive_ptr<TextRange> textRange2 = new TextRange();
	para0->GetTextRanges()->Append(textRange1);
	para0->GetTextRanges()->Append(textRange2);

	intrusive_ptr<TextParagraph> para1 = new TextParagraph();
	tf->GetParagraphs()->Append(para1);
	intrusive_ptr<TextRange> textRange11 = new TextRange();
	intrusive_ptr<TextRange> textRange12 = new TextRange();
	intrusive_ptr<TextRange> textRange13 = new TextRange();
	para1->GetTextRanges()->Append(textRange11);
	para1->GetTextRanges()->Append(textRange12);
	para1->GetTextRanges()->Append(textRange13);

	intrusive_ptr<TextParagraph> para2 = new TextParagraph();
	tf->GetParagraphs()->Append(para2);
	intrusive_ptr<TextRange> textRange21 = new TextRange();
	intrusive_ptr<TextRange> textRange22 = new TextRange();
	intrusive_ptr<TextRange> textRange23 = new TextRange();
	para2->GetTextRanges()->Append(textRange21);
	para2->GetTextRanges()->Append(textRange22);
	para2->GetTextRanges()->Append(textRange23);

	for (int i = 0; i < 3; i++)
	{
		for (int j = 0; j < 3; j++)
		{
			tf->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(j)->SetText((L"TextRange " + to_wstring(j)).c_str());
			if (j == 0)
			{
				tf->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(j)->GetFill()->SetFillType(FillFormatType::Solid);
				tf->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(j)->GetFill()->GetSolidColor()->SetColor(Color::GetLightBlue());
				tf->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(j)->GetFormat()->SetIsBold(TriState::True);
				tf->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(j)->SetFontHeight(15);
			}
			else if (j == 1)
			{
				tf->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(j)->GetFill()->SetFillType(FillFormatType::Solid);
				tf->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(j)->GetFill()->GetSolidColor()->SetColor(Color::GetBlue());
				tf->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(j)->GetFormat()->SetIsItalic(TriState::True);
				tf->GetParagraphs()->GetItem(i)->GetTextRanges()->GetItem(j)->SetFontHeight(18);
			}
		}
	}
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

