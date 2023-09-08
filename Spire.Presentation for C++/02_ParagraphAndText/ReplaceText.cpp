#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"TextTemplate.pptx";
	wstring outputFile = OUTPUTPATH"ReplaceText.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	slide->ReplaceAllText(L"Spire.Presentation for C++", L"Spire.PPT", false);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

