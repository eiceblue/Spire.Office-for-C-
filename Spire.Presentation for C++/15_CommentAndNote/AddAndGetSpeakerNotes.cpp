#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_1.pptx";
	wstring outputFile_px = OUTPUTPATH"AddAndGetSpeakerNotes.pptx";
	wstring outputFile_txt = OUTPUTPATH"AddAndGetSpeakerNotes.txt";

	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide and in the PowerPoint document.
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Get the NotesSlide in the first slide,if there is no notes, we need to add it firstly.
	intrusive_ptr<NotesSlide> ns = slide->GetNotesSlide();
	if (ns == nullptr)
	{
		ns = slide->AddNotesSlide();
	}
	//Add the text string as the notes.
	ns->GetNotesTextFrame()->SetText(L"Speak notes added by Spire.Presentation");
	wofstream desFile(outputFile_txt, ios::out);
	//Get the speaker notes and save to txt file.
	desFile << "The speaker notes added by Spire.Presentation is: " << ns->GetNotesTextFrame()->GetText() << endl;
	desFile.close();

	//Save to PowerPoint file.
	presentation->SaveToFile(outputFile_px.c_str(), FileFormat::Pptx2013);
}
