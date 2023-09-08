#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddSmartArtNode.pptx";
	wstring outputFile = OUTPUTPATH"AssistantNode.pptx";


	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());
	intrusive_ptr<ISmartArtNode> node;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Get the SmartArt and collect nodes
		intrusive_ptr<ISmartArt> smartArt = Object::Dynamic_cast<ISmartArt >(shape);
		if (smartArt != nullptr)
		{
			intrusive_ptr<ISmartArtNodeCollection> nodes = smartArt->GetNodes();

			//Traverse through all nodes inside SmartArt
			for (int i = 0; i < nodes->GetCount(); i++)
			{
				//Access SmartArt node at index i
				node = nodes->GetItem(i);
				// Check if node is assitant node
				if (!node->GetIsAssistant())
				{
					//Set node as assitant node
					node->SetIsAssistant(true);
				}
			}
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
			
}
