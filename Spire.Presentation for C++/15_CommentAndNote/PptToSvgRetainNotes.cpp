#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_5.pptx";
	wstring outputFile = OUTPUTPATH"SVG/PptToSvgRetainNotes/";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Retain the notes while converting PowerPoint file to svg file.
	ppt->SetIsNoteRetained(true);
	intrusive_ptr<SlideCollection> slides = ppt->GetSlides();
	for (int i = 0; i < slides->GetCount(); i++)
	{
		intrusive_ptr<Stream> svg = slides->GetItem(i)->SaveToSVG();
		svg->Save((outputFile + L"output_" + to_wstring(i) + L".svg").c_str());
		svg->Dispose();
	}
}
