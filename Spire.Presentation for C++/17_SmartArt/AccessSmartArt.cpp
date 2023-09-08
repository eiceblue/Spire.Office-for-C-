#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SmartArt.pptx";
	wstring outputFile = OUTPUTPATH"AccessSmartArt.txt";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	wstring* content = new std::wstring();

	content->append(L"Access SmartArt nodes.");
	content->append(L"\r\nHere is the SmartArt node parameters details:");
	intrusive_ptr<ISmartArtNode> node;
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Get the SmartArt
		intrusive_ptr<ISmartArt> sa = Object::Dynamic_cast<ISmartArt>(shape);
		if (sa != nullptr)
		{
			intrusive_ptr<ISmartArtNodeCollection> nodes = sa->GetNodes();

			//Traverse through all nodes inside SmartArt
			for (int i = 0; i < nodes->GetCount(); i++)
			{
				//Access SmartArt node at index i
				node = nodes->GetItem(i);
				std::wstring text = node->GetTextFrame()->GetText();
				//Print the SmartArt node parameters
				content->append(L"\r\nNode text = " + text + L", Node level = "+ std::to_wstring(node->GetLevel()) + L", Node Position = " + std::to_wstring(node->GetPosition()));
			}
		}
	}
	//Save the file
	wofstream write(outputFile);
	write << content->c_str();
	write.close();
	
}
