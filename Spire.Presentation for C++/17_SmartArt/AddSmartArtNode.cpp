#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddSmartArtNode.pptx";
	wstring outputFile = OUTPUTPATH"AddSmartArtNode.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the SmartArt
	intrusive_ptr<ISmartArt> sa = Object::Dynamic_cast<ISmartArt>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Add a node
	intrusive_ptr<ISmartArtNode> node = sa->GetNodes()->AddNode();
	//Add text and set the text style 
	node->GetTextFrame()->SetText(L"AddText");
	node->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
	node->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::HotPink);

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
				
}
