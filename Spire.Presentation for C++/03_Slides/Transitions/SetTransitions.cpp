#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SetTransitions.pptx";
	wstring outputFile = OUTPUTPATH"SetTransitions.pptx";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT with password
	presentation->LoadFromFile(inputFile.c_str());

	//Set the first slide transition as push and sound mode
	presentation->GetSlides()->GetItem(0)->GetSlideShowTransition()->SetType(TransitionType::Push);
	presentation->GetSlides()->GetItem(0)->GetSlideShowTransition()->SetSoundMode(TransitionSoundMode::StartSound);

	//Set the second slide transition as circle and set the speed 
	presentation->GetSlides()->GetItem(1)->GetSlideShowTransition()->SetType(TransitionType::Fade);
	presentation->GetSlides()->GetItem(1)->GetSlideShowTransition()->SetSpeed(TransitionSpeed::Slow);

	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	
}
