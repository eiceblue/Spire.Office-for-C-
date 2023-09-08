#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddSmartArtNode.pptx";
	wstring outputFile = OUTPUTPATH"ChangeNodeText.pptx";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Get the SmartArt and collect nodes
		intrusive_ptr<ISmartArt> smartArt = Object::Dynamic_cast<ISmartArt> (shape);
		if (smartArt != nullptr)
		{
			//Obtain the reference of a node by using its Index  
			// select second root node
			intrusive_ptr<ISmartArtNode> node = smartArt->GetNodes()->GetItem(1);
			// Set the text of the TextFrame 
			node->GetTextFrame()->SetText(L"Second root node");
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
				
}
