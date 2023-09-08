#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"Hyperlinks.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background Image
	std::wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);

	//Add new shape to PPT document
	intrusive_ptr<RectangleF> rec = new RectangleF(ppt->GetSlideSize()->GetSize()->GetWidth() / 2 - 255, 120, 500, 280);
	intrusive_ptr<IAutoShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec);
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetLine()->SetWidth(0);

	//Add some paragraphs with hyperlinks
	intrusive_ptr<TextParagraph> para1 = new TextParagraph();
	intrusive_ptr<TextRange> tr = new TextRange(L"E-iceblue");
	tr->GetFill()->SetFillType(FillFormatType::Solid);
	tr->GetFill()->GetSolidColor()->SetColor(Color::GetBlue());
	para1->GetTextRanges()->Append(tr);
	para1->SetAlignment(TextAlignmentType::Center);
	shape->GetTextFrame()->GetParagraphs()->Append(para1);
	intrusive_ptr<TextParagraph> tp = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(tp);

	//Add some paragraphs with hyperlinks
	intrusive_ptr<TextParagraph> para2 = new TextParagraph();
	intrusive_ptr<TextRange> tr1 = new TextRange(L"Click to know more about Spire.Presentation.");
	tr1->GetClickAction()->SetAddress(L"https://www.e-iceblue.com/Introduce/presentation-for-CPP.html");
	para2->GetTextRanges()->Append(tr1);
	shape->GetTextFrame()->GetParagraphs()->Append(para2);
	intrusive_ptr<TextParagraph> tp2 = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(tp2);

	intrusive_ptr<TextParagraph> para3 = new TextParagraph();
	intrusive_ptr<TextRange> tr2 = new TextRange(L"Click to visit E-iceblue Home page.");
	tr2->GetClickAction()->SetAddress(L"https://www.e-iceblue.com/");
	para3->GetTextRanges()->Append(tr2);
	shape->GetTextFrame()->GetParagraphs()->Append(para3);
	intrusive_ptr<TextParagraph> temp3 = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(temp3);

	intrusive_ptr<TextParagraph> para4 = new TextParagraph();
	intrusive_ptr<TextRange> tr3 = new TextRange(L"Click to go to the forum to raise questions.");
	tr3->GetClickAction()->SetAddress(L"https://www.e-iceblue.com/forum/components-f5.html");
	para4->GetTextRanges()->Append(tr3);
	shape->GetTextFrame()->GetParagraphs()->Append(para4);
	intrusive_ptr<TextParagraph> temp4 = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(temp4);

	intrusive_ptr<TextParagraph> para5 = new TextParagraph();
	intrusive_ptr<TextRange> tr4 = new TextRange(L"Click to contact our sales team via email.");
	tr4->GetClickAction()->SetAddress(L"mailto:sales@e-iceblue.com");
	para5->GetTextRanges()->Append(tr4);
	shape->GetTextFrame()->GetParagraphs()->Append(para5);
	intrusive_ptr<TextParagraph> temp5 = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(temp5);

	intrusive_ptr<TextParagraph> para6 = new TextParagraph();
	intrusive_ptr<TextRange> tr5 = new TextRange(L"Click to contact our support team via email.");
	tr5->GetClickAction()->SetAddress(L"mailto:support@e-iceblue.com");
	para6->GetTextRanges()->Append(tr5);
	shape->GetTextFrame()->GetParagraphs()->Append(para6);

	for (int i = 0; i < shape->GetTextFrame()->GetParagraphs()->GetCount(); i++)
	{
		intrusive_ptr<TextParagraph> temp6 = shape->GetTextFrame()->GetParagraphs()->GetItem(i);
		if (temp6->GetTextRanges()->GetCount() > 0)
		{
			temp6->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));
			temp6->GetTextRanges()->GetItem(0)->SetFontHeight(20);
		}
	}

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
