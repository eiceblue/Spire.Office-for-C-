#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_4.pptx";
	wstring outputFile = OUTPUTPATH"CheckPasswordProtection.txt";
	//Create Presentation
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Check whether a PPT document is password protected
	bool isProtected = presentation->IsPasswordProtected(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);
	outFile << "The file is " << (isProtected ? "password " : "not password ") << "protected!" << endl;

	//Save the file
	outFile.close();
}
