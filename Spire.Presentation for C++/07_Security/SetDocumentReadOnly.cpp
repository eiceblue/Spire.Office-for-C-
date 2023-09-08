#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SetDocumentReadOnly.pptx";
	wstring outputFile = OUTPUTPATH"SetDocumentReadOnly.pptx";
	
	//Load a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the password that the user entered
	std::wstring password = L"e-iceblue";

	//Protect the document with the password
	presentation->Protect(password.c_str());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

}

