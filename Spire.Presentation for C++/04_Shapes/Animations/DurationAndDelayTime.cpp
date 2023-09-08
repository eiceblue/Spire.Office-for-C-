#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Animation.pptx";
	wstring outputFile = OUTPUTPATH"DurationAndDelayTime.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> presentation = new Presentation();

	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);
	intrusive_ptr<AnimationEffectCollection> animations = slide->GetTimeline()->GetMainSequence();

	//Get duration time of animation
	float durationTime = animations->GetItem(0)->GetTiming()->GetDuration();

	//Set new duration time of animation
	animations->GetItem(0)->GetTiming()->SetDuration(0.8f);

	//Get delay time of animation
	float delayTime = animations->GetItem(0)->GetTiming()->GetTriggerDelayTime();

	//Set new delay time of animation
	animations->GetItem(0)->GetTiming()->SetTriggerDelayTime(0.6f);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

