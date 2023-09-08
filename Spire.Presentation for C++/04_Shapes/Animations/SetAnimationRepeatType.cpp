#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Animation.pptx";
	wstring outputFile = OUTPUTPATH"SetAnimationRepeatType.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> presentation = new Presentation();
	//Load file
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);
	intrusive_ptr<AnimationEffectCollection> animations = slide->GetTimeline()->GetMainSequence();
	animations->GetItem(0)->GetTiming()->SetAnimationRepeatType(AnimationRepeatType::UtilEndOfSlide);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

