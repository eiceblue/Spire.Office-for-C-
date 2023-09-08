#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Alignment.pptx";
	wstring outputFile = OUTPUTPATH"Alignment.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	ppt->LoadFromFile(inputFile.c_str());

	//Get the related shape and set the text alignment
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Left);
	shape->GetTextFrame()->GetParagraphs()->GetItem(1)->SetAlignment(TextAlignmentType::Center);
	shape->GetTextFrame()->GetParagraphs()->GetItem(2)->SetAlignment(TextAlignmentType::Right);
	shape->GetTextFrame()->GetParagraphs()->GetItem(3)->SetAlignment(TextAlignmentType::Justify);
	shape->GetTextFrame()->GetParagraphs()->GetItem(4)->SetAlignment(TextAlignmentType::None);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}

