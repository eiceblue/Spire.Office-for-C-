#include "pch.h"

using namespace Spire::Common;
using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"EmbedSoundFile.pdf";
	wstring inputFile = DATAPATH"EmbedSoundFile.pdf";

	//Create a new PDF document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	doc->LoadFromFile(inputFile.c_str());

	//Add a page
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	//Create a sound action
	wstring inputFile_wav = DATAPATH"Music.wav";
	intrusive_ptr<PdfSoundAction> soundAction = new PdfSoundAction(inputFile_wav.c_str());
	soundAction->GetSound()->SetBits(15);
	soundAction->GetSound()->SetChannels(PdfSoundChannels::Stereo);
	soundAction->GetSound()->SetEncoding(PdfSoundEncoding::Signed);
	soundAction->SetVolume(0.8f);
	soundAction->SetRepeat(true);

	// Set the sound action to be executed when the PDF document is opened
	doc->SetAfterOpenAction(soundAction);
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}