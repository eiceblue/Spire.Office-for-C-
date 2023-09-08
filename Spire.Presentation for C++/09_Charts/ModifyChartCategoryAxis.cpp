
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChartSample2.pptx";
	wstring outputFile = OUTPUTPATH"ModifyChartCategoryAxis.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Modify the major unit
	chart->GetPrimaryCategoryAxis()->SetIsAutoMajor(false);
	chart->GetPrimaryCategoryAxis()->SetMajorUnit(1);
	chart->GetPrimaryCategoryAxis()->SetMajorUnitScale(ChartBaseUnitType::Months);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
