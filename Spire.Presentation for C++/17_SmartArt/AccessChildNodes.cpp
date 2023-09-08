#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SmartArt.pptx";
	wstring outputFile = OUTPUTPATH"AccessChildNode.txt";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	wstring* content = new std::wstring();

	content->append(L"Access SmartArt child nodes.");
	content->append(L"\r\nHere is the SmartArt child node parameters details:");
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Get the SmartArt and collect nodes
		intrusive_ptr<ISmartArt> sa = Object::Dynamic_cast<ISmartArt>(shape);
		if (sa != nullptr)
		{
			intrusive_ptr<ISmartArtNodeCollection> nodes = sa->GetNodes();
			int position = 0;
			//Access the parent node at position 0
			intrusive_ptr<ISmartArtNode> node = nodes->GetItem(position);
			intrusive_ptr<ISmartArtNode> childnode;
			//Traverse through all child nodes inside SmartArt
			for (int i = 0; i < node->GetChildNodes()->GetCount(); i++)
			{
				//Access SmartArt child node at index i
				childnode = node->GetChildNodes()->GetItem(i);
				std::wstring text = childnode->GetTextFrame()->GetText();
				//Print the SmartArt child node parameters       
				content->append(L"\r\nNode text = " + text + L", Node level = "+ std::to_wstring(childnode->GetLevel()) +L", Node Position = "+ std::to_wstring(childnode->GetPosition()));
			}
		}
	}
	//Save the file
	wofstream write(outputFile);
	write << content->c_str();
	write.close();
	
}
