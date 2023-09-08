#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Bullets2.pptx";
	wstring outputFile = OUTPUTPATH"MultipleLevelBullets.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Access the first placeholder in the slide and typecasting it as AutoShape
	intrusive_ptr<ITextFrameProperties> tf1 = (Object::Dynamic_cast<IAutoShape>(slide->GetShapes()->GetItem(1)))->GetTextFrame();

	//Access the first Paragraph and set bullet style
	intrusive_ptr<TextParagraph> para = tf1->GetParagraphs()->GetItem(0);
	para->SetBulletType(TextBulletType::Symbol);
	para->SetBulletChar(8226);
	para->SetDepth(0);

	//Access the second Paragraph and set bullet style
	para = tf1->GetParagraphs()->GetItem(1);
	para->SetBulletType(TextBulletType::Symbol);
	para->SetBulletChar(45);
	para->SetDepth(1);

	//Access the third Paragraph and set bullet style
	para = tf1->GetParagraphs()->GetItem(2);
	para->SetBulletType(TextBulletType::Symbol);
	//para->SetBulletChar(dynamic_cast<wchar_t>(8226));
	para->SetBulletChar(8226);
	para->SetDepth(2);

	//Access the fourth Paragraph and set bullet style
	para = tf1->GetParagraphs()->GetItem(3);
	para->SetBulletType(TextBulletType::Symbol);
	para->SetBulletChar(45);
	para->SetDepth(3);

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}

