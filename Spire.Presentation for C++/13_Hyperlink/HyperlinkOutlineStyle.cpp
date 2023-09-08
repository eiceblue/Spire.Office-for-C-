#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"HyperlinkOutlineStyle.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Add new shape to PPT document
	intrusive_ptr<RectangleF> rec = new RectangleF(ppt->GetSlideSize()->GetSize()->GetWidth() / 2 - 255, 120, 400, 100);
	intrusive_ptr<IAutoShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rec);
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetLine()->SetFillType(FillFormatType::None);

	//Add a paragraph with hyperlink
	intrusive_ptr<TextParagraph> para1 = new TextParagraph();
	intrusive_ptr<TextRange> tr1 = new TextRange(L"Click to know more about Spire.Presentation");
	tr1->GetClickAction()->SetAddress(L"https://www.e-iceblue.com/Introduce/presentation-for-CPP.html");
	para1->GetTextRanges()->Append(tr1);

	//Set the format of textrange
	tr1->GetFormat()->SetFontHeight(20.0);
	tr1->SetIsItalic(TriState::True);

	//Set the outline format of textrange
	tr1->GetTextLineFormat()->GetFillFormat()->SetFillType(FillFormatType::Solid);
	tr1->GetTextLineFormat()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetLightSeaGreen());
	tr1->GetTextLineFormat()->SetJoinStyle(LineJoinType::Round);
	tr1->GetTextLineFormat()->SetWidth(2.0);

	//Add the paragraph to shape
	shape->GetTextFrame()->GetParagraphs()->Append(para1);
	intrusive_ptr<TextParagraph> tp = new TextParagraph();
	shape->GetTextFrame()->GetParagraphs()->Append(tp);

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
