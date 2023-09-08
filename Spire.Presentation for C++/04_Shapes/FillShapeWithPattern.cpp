#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"FillShapeWithGradient.pptx";
	wstring outputFile = OUTPUTPATH"FillShapeWithGradient.pptx";

	//Load a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	ppt->LoadFromFile(inputFile.c_str());

	//Get the first shape and set the style to be Gradient
	intrusive_ptr<IAutoShape> GradientShape = Object::Dynamic_cast<IAutoShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	GradientShape->GetFill()->SetFillType(FillFormatType::Gradient);
	GradientShape->GetFill()->GetGradient()->GetGradientStops()->Append(0, Color::GetLightSkyBlue());
	GradientShape->GetFill()->GetGradient()->GetGradientStops()->Append(1, Color::GetLightGray());

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}

