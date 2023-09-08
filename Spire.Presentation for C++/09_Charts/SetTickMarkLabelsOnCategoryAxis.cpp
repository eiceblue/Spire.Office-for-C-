#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_3.pptx";
	wstring outputFile = OUTPUTPATH"SetTickMarkLabelsOnCategoryAxis.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Rotate tick labels.
	chart->GetPrimaryCategoryAxis()->SetTextRotationAngle(45);

	//Specify interval between labels.
	chart->GetPrimaryCategoryAxis()->SetIsAutomaticTickLabelSpacing(false);
	chart->GetPrimaryCategoryAxis()->SetTickLabelSpacing(2);

	//Change position.
	chart->GetPrimaryCategoryAxis()->SetTickLabelPosition(TickLabelPositionType::TickLabelPositionHigh);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
