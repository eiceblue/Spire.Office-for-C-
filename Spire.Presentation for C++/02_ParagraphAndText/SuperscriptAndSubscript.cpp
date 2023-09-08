#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Az.pptx";
	wstring outputFile = OUTPUTPATH"SuperscriptAndSubscript.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);
	//Add a shape 
	intrusive_ptr<IAutoShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(150, 100, 200, 50));
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetTextFrame()->GetParagraphs()->Clear();

	shape->AppendTextFrame(L"Test");
	intrusive_ptr<TextRange> tr = new TextRange(L"superscript");
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->Append(tr);

	//Set superscript text
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(1)->GetFormat()->SetScriptDistance(30);

	intrusive_ptr<TextRange> textRange = shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0);
	textRange->GetFill()->SetFillType(FillFormatType::Solid);
	textRange->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());
	textRange->SetFontHeight(20);
	textRange->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));

	textRange = shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(1);
	textRange->GetFill()->SetFillType(FillFormatType::Solid);
	textRange->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());
	textRange->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));


	shape = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(150, 150, 200, 50));
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetTextFrame()->GetParagraphs()->Clear();

	shape->AppendTextFrame(L"Test");
	tr = new TextRange(L"subscript");
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->Append(tr);

	//Set subscript text
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(1)->GetFormat()->SetScriptDistance(-25);

	textRange = shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0);
	textRange->GetFill()->SetFillType(FillFormatType::Solid);
	textRange->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());
	textRange->SetFontHeight(20);
	textRange->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));

	textRange = shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(1);
	textRange->GetFill()->SetFillType(FillFormatType::Solid);
	textRange->GetFill()->GetSolidColor()->SetColor(Color::GetCadetBlue());
	textRange->SetLatinFont(new TextFont(L"Lucida Sans Unicode"));

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

