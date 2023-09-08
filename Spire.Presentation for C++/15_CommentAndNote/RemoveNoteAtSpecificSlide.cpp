#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"RemoveNoteFromSlides.pptx";
	wstring outputFile = OUTPUTPATH"RemoveNotesAtSpecificSlide.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Get note slide
	intrusive_ptr<NotesSlide> note = slide->GetNotesSlide();
	//Clear note text
	note->GetNotesTextFrame()->SetText(L"");

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
