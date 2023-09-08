#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddSmartArtNode2.pptx";
	wstring outputFile = OUTPUTPATH"AddNodeByPosition.pptx";


	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Get the SmartArt and collect nodes
		intrusive_ptr<ISmartArt> smartArt = Object::Dynamic_cast<ISmartArt >(shape);
		if (smartArt != nullptr)
		{
			int position = 0;
			//Add a new node at specific position
			intrusive_ptr<ISmartArtNode> node = smartArt->GetNodes()->AddNodeByPosition(position);
			//Add text and set the text style 
			node->GetTextFrame()->SetText(L"New Node");
			node->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
			node->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Red);

			//Get a node
			node = smartArt->GetNodes()->GetItem(1);
			position = 1;
			//Add a new child node at specific position
			intrusive_ptr<ISmartArtNode> childNode = node->GetChildNodes()->AddNodeByPosition(position);
			//Add text and set the text style 
			node->GetTextFrame()->SetText(L"New child node");
			node->GetTextFrame()->GetTextRange()->GetFill()->SetFillType(FillFormatType::Solid);
			node->GetTextFrame()->GetTextRange()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Blue);
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
			
}
