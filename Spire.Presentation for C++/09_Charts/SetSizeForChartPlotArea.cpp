#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChartSample2.pptx";
	wstring outputFile = OUTPUTPATH"SetSizeForChartPlotArea.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Set width and height for chart plot area
	chart->GetPlotArea()->SetWidth(250);
	chart->GetPlotArea()->SetHeight(300);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
