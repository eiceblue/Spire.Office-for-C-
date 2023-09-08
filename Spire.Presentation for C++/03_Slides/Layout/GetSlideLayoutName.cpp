#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"GetSlideLayoutName.pptx";
	wstring outputFile = OUTPUTPATH"GetSlideLayoutName.txt";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);

	//Loop through the slides of PPT document
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		//Get the name of slide layout
		std::wstring name = presentation->GetSlides()->GetItem(i)->GetLayout()->GetName();
		outFile << "The name of slide " << i << " layout is :" << name.c_str() << endl;
	}
	outFile.close();


}
