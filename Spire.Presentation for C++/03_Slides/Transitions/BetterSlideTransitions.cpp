#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SetTransitions.pptx";
	wstring outputFile = OUTPUTPATH"BetterSlideTransitions.pptx";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT
	presentation->LoadFromFile(inputFile.c_str());

	//Set the first slide transition as circle
	presentation->GetSlides()->GetItem(0)->GetSlideShowTransition()->SetType(TransitionType::Circle);

	// Set the transition time of 3 seconds
	presentation->GetSlides()->GetItem(0)->GetSlideShowTransition()->SetAdvanceOnClick(true);
	presentation->GetSlides()->GetItem(0)->GetSlideShowTransition()->SetAdvanceAfterTime(3000);

	//Set the second slide transition as comb and set the speed 
	presentation->GetSlides()->GetItem(1)->GetSlideShowTransition()->SetType(TransitionType::Comb);
	presentation->GetSlides()->GetItem(1)->GetSlideShowTransition()->SetSpeed(TransitionSpeed::Slow);

	// Set the transition time of 5 seconds
	presentation->GetSlides()->GetItem(1)->GetSlideShowTransition()->SetAdvanceOnClick(true);
	presentation->GetSlides()->GetItem(1)->GetSlideShowTransition()->SetAdvanceAfterTime(5000);

	// Set the third slide transition as zoom
	presentation->GetSlides()->GetItem(2)->GetSlideShowTransition()->SetType(TransitionType::Zoom);

	// Set the transition time of 7 seconds
	presentation->GetSlides()->GetItem(2)->GetSlideShowTransition()->SetAdvanceOnClick(true);
	presentation->GetSlides()->GetItem(2)->GetSlideShowTransition()->SetAdvanceAfterTime(7000);

	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	
}
