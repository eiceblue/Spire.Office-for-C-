
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"GroupTwoLevelAxisLabels.pptx";
	wstring outputFile = OUTPUTPATH"GroupTwoLevelAxisLabels.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Get the category axis from the chart.
	intrusive_ptr<IChartAxis> chartAxis = chart->GetPrimaryCategoryAxis();

	//Group the axis labels that have the same first-level label.
	if (chartAxis->GetHasMultiLvlLbl())
	{
		chartAxis->SetIsMergeSameLabel(true);
	}

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
