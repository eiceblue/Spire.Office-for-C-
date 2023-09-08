#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	
	wstring inputFile = DATAPATH"Indent.pptx";
	wstring outputFile = OUTPUTPATH"Indent.pptx";

	//Load a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	intrusive_ptr<ParagraphCollection> paras = shape->GetTextFrame()->GetParagraphs();

	//Set the paragraph style for first paragraph
	paras->GetItem(0)->SetIndent(20);
	paras->GetItem(0)->SetLeftMargin(10);
	paras->GetItem(0)->SetSpaceAfter(10);

	//Set the paragraph style of the third paragraph 
	paras->GetItem(2)->SetIndent(-100);
	paras->GetItem(2)->SetLeftMargin(40);
	paras->GetItem(2)->SetSpaceBefore(0);
	paras->GetItem(2)->SetSpaceAfter(0);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

}