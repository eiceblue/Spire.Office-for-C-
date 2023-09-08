#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Encrypt.pptx";
	wstring outputFile = OUTPUTPATH"Encrypt.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the password that the user entered
	std::wstring password = L"e-iceblue";

	//Encrypy the document with the password
	presentation->Encrypt(password.c_str());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}