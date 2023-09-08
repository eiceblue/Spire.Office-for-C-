#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"RemoveNode.pptx";
	wstring outputFile = OUTPUTPATH"RemoveNode.pptx";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the SmartArt and collect nodes
	intrusive_ptr<ISmartArt> sa = Object::Dynamic_cast<ISmartArt>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	intrusive_ptr<ISmartArtNodeCollection> nodes = sa->GetNodes();

	//Remove the node to specific position
	nodes->RemoveNodeByPosition(2);

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
			
}
