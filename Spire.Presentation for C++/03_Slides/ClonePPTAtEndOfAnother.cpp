#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	
	wstring inputFile_1 = DATAPATH"ChangeSlidePosition.pptx";
	wstring inputFile_2 = DATAPATH"PPTSample_N.pptx";
	wstring outputFile = OUTPUTPATH"ClonePPTAtEndOfAnother.pptx";

	//Load source document from disk
	intrusive_ptr<Presentation> sourcePPT = new Presentation();
	sourcePPT->LoadFromFile(inputFile_1.c_str());

	//Load destination document from disk
	intrusive_ptr<Presentation> destPPT = new Presentation();
	destPPT->LoadFromFile(inputFile_2.c_str());

	//Loop through all slides of source document
	for (int l = 0; l < sourcePPT->GetSlides()->GetCount(); l++)
	{
		intrusive_ptr<ISlide> slide = sourcePPT->GetSlides()->GetItem(l);
		//Append the slide at the end of destination document
		destPPT->GetSlides()->Append(slide);
	}
	//Save the document
	destPPT->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
		
}
