#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile_1 = DATAPATH"TextTemplate.pptx";
	wstring inputFile_2 = DATAPATH"CopyParagraph.pptx";
	wstring outputFile = OUTPUTPATH"CopyParagraphToAnotherPPT.pptx";

	//Load the source file
	intrusive_ptr<Presentation> ppt1 = new Presentation();
	ppt1->LoadFromFile(inputFile_1.c_str());

	//Get the text from the first shape on the first slide
	intrusive_ptr<IShape> sourceshp = ppt1->GetSlides()->GetItem(0)->GetShapes()->GetItem(0);
	wstring text = (Object::Dynamic_cast<IAutoShape>(sourceshp))->GetTextFrame()->GetText();

	//Load the target file
	intrusive_ptr<Presentation> ppt2 = new Presentation();
	ppt2->LoadFromFile(inputFile_2.c_str());

	//Get the first shape on the first slide from the target file
	intrusive_ptr<IShape> destshp = ppt2->GetSlides()->GetItem(0)->GetShapes()->GetItem(0);

	//Add the text to the target file
	wstring text2 = (Object::Dynamic_cast<IAutoShape>(destshp))->GetTextFrame()->GetText();
	(Object::Dynamic_cast<IAutoShape>(destshp))->GetTextFrame()->SetText((text2 + L"\n\n" + text).c_str());

	//Save the document
	ppt2->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);


}

