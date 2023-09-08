#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Az2.pptx";
	wstring outputFile = OUTPUTPATH"SetParagraphFont.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Access the first and second placeholder in the slide and typecasting it as AutoShape
	intrusive_ptr<ITextFrameProperties> tf1 = (Object::Dynamic_cast<IAutoShape>(slide->GetShapes()->GetItem(0)))->GetTextFrame();
	intrusive_ptr<ITextFrameProperties> tf2 = (Object::Dynamic_cast<IAutoShape>(slide->GetShapes()->GetItem(1)))->GetTextFrame();

	// Access the first Paragraph
	intrusive_ptr<TextParagraph> para1 = tf1->GetParagraphs()->GetItem(0);
	intrusive_ptr<TextParagraph> para2 = tf2->GetParagraphs()->GetItem(0);

	//Justify the paragraph
	para2->SetAlignment(TextAlignmentType::Justify);

	//Access the first text range
	intrusive_ptr<TextRange> textRange1 = para1->GetFirstTextRange();
	intrusive_ptr<TextRange> textRange2 = para2->GetFirstTextRange();

	//Define new fonts
	intrusive_ptr<TextFont> fd1 = new TextFont(L"Elephant");
	intrusive_ptr<TextFont> fd2 = new TextFont(L"Castellar");

	// Assign new fonts to text range
	textRange1->SetLatinFont(fd1);
	textRange2->SetLatinFont(fd2);

	// Set font to Bold
	textRange1->GetFormat()->SetIsBold(TriState::True);
	textRange2->GetFormat()->SetIsBold(TriState::False);

	// Set font to Italic
	textRange1->GetFormat()->SetIsItalic(TriState::False);
	textRange2->GetFormat()->SetIsItalic(TriState::True);

	// Set font color
	textRange1->GetFill()->SetFillType(FillFormatType::Solid);
	textRange1->GetFill()->GetSolidColor()->SetColor(Color::GetPurple());
	textRange2->GetFill()->SetFillType(FillFormatType::Solid);
	textRange2->GetFill()->GetSolidColor()->SetColor(Color::GetPeru());

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

