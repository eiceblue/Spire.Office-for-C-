
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"PPTSample_N.pptx";
	wstring outputFile = OUTPUTPATH"SetGradientBackground.pptx";

	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Set the background to gradient
	slide->GetSlideBackground()->SetType(BackgroundType::Custom);
	slide->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Gradient);

	//Add gradient stops
	slide->GetSlideBackground()->GetFill()->GetGradient()->GetGradientStops()->Append(0.1f, KnownColors::LightSeaGreen);
	slide->GetSlideBackground()->GetFill()->GetGradient()->GetGradientStops()->Append(0.7f, KnownColors::LightCyan);

	//Set gradient shape type
	slide->GetSlideBackground()->GetFill()->GetGradient()->SetGradientShape(GradientShapeType::Linear);

	//Set the angle
	slide->GetSlideBackground()->GetFill()->GetGradient()->GetLinearGradientFill()->SetAngle(45);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
