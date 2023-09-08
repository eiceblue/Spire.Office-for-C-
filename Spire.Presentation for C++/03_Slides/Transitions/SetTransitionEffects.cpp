#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SetTransitions.pptx";
	wstring outputFile = OUTPUTPATH"SetTransitionEffects.pptx";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	// Set effects
	presentation->GetSlides()->GetItem(0)->GetSlideShowTransition()->SetType(TransitionType::Cut);
	(Object::Dynamic_cast<OptionalBlackTransition>(presentation->GetSlides()->GetItem(0)->GetSlideShowTransition()->GetValue()))->SetFromBlack(true);

	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

	
	
}
