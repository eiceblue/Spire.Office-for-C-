#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChangeTextStyle.pptx";
	wstring outputFile = OUTPUTPATH"ChangeTextStyle.pptx";

	//Load a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	intrusive_ptr<ParagraphCollection> paras = shape->GetTextFrame()->GetParagraphs();

	//Set the style for the text content in the first paragraph
	for (int t = 0; t < paras->GetItem(0)->GetTextRanges()->GetCount(); t++)
	{
		intrusive_ptr<TextRange> tr = paras->GetItem(0)->GetTextRanges()->GetItem(t);
		tr->GetFill()->SetFillType(FillFormatType::Solid);
		tr->GetFill()->GetSolidColor()->SetColor(Color::GetForestGreen());
		tr->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));
		tr->SetFontHeight(14);
	}
	//Set the style for the text content in the third paragraph
	for (int t = 0; t < paras->GetItem(2)->GetTextRanges()->GetCount(); t++)
	{
		intrusive_ptr<TextRange> tr = paras->GetItem(2)->GetTextRanges()->GetItem(t);
		//tr->GetFill()->SetFillType(Spire::Presentation::Drawing::FillFormatType::Solid);
		tr->GetFill()->SetFillType(FillFormatType::Solid);
		tr->GetFill()->GetSolidColor()->SetColor(Color::GetCornflowerBlue());
		tr->SetLatinFont(new TextFont(L"Calibri"));
		tr->SetFontHeight(16);
		tr->SetTextUnderlineType(TextUnderlineType::Dashed);
	}

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2007);

}

