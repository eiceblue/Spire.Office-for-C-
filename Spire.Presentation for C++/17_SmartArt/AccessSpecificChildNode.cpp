#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SmartArt.pptx";
	wstring outputFile = OUTPUTPATH"AccessSpecificChildNode.txt";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);
	outFile << "Access SmartArt child node at specific position." << endl;
	outFile << "Here is the SmartArt child node parameters details:" << endl;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Get the SmartArt
		intrusive_ptr<ISmartArt> sa = Object::Dynamic_cast<ISmartArt>(shape);
		if (sa != nullptr)
		{
			//Get SmartArt node collection 
			intrusive_ptr<ISmartArtNodeCollection> nodes = sa->GetNodes();

			//Access SmartArt node at index 0
			intrusive_ptr<ISmartArtNode> node = nodes->GetItem(0);

			//Access SmartArt child node at index 1
			intrusive_ptr<ISmartArtNode> childNode = node->GetChildNodes()->GetItem(1);

			//Print the SmartArt child node parameters
			outFile << "Node text = " << childNode->GetTextFrame()->GetText() << ", Node level = " << childNode->GetLevel() << ", Node Position = " << childNode->GetPosition() << endl;
		}
	}
	//Save the file
	outFile.close();
	
}
