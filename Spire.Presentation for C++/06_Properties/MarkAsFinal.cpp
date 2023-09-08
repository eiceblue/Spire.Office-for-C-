#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"MarkAsFinal.pptx";
	wstring outputFile = OUTPUTPATH"MarkAsFinal.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Mark the document as final
	//presentation->GetDocumentProperty()->SetItem(L"_MarkAsFinal","true");

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

}

