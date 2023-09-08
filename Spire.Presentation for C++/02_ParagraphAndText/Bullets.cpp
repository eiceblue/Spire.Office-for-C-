#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Bullets.pptx";
	wstring outputFile = OUTPUTPATH"Bullets.pptx";

	//Load a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));

	for (int t = 0; t < shape->GetTextFrame()->GetParagraphs()->GetCount(); t++)
	{
		intrusive_ptr<TextParagraph> para = shape->GetTextFrame()->GetParagraphs()->GetItem(t);
		//Add the bullets
		para->SetBulletType(TextBulletType::Numbered);
		para->SetBulletStyle(NumberedBulletStyle::BulletRomanLCPeriod);
	}

	//Save the document and launch
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);


}
