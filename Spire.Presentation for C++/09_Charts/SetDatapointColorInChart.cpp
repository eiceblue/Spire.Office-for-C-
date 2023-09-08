#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SetDatapointColorInChart.pptx";
	wstring outputFile = OUTPUTPATH"SetDatapointColorInChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Initialize an instances of dataPoint
	intrusive_ptr<ChartDataPoint> cdp1 = new ChartDataPoint(chart->GetSeries()->GetItem(0));

	//Specify the datapoint order
	cdp1->SetIndex(0);

	//Set the color of the datapoint
	cdp1->GetFill()->SetFillType(FillFormatType::Solid);
	cdp1->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Orange);

	//Add the dataPoint to first series
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(cdp1);

	//Set the color for the other three data points
	intrusive_ptr<ChartDataPoint> cdp2 = new ChartDataPoint(chart->GetSeries()->GetItem(0));
	cdp2->SetIndex(1);
	cdp2->GetFill()->SetFillType(FillFormatType::Solid);
	cdp2->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Gold);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(cdp2);

	intrusive_ptr<ChartDataPoint> cdp3 = new ChartDataPoint(chart->GetSeries()->GetItem(0));
	cdp3->SetIndex(2);
	cdp3->GetFill()->SetFillType(FillFormatType::Solid);
	cdp3->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::MediumPurple);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(cdp3);

	intrusive_ptr<ChartDataPoint> cdp4 = new ChartDataPoint(chart->GetSeries()->GetItem(0));
	cdp4->SetIndex(1);
	cdp4->GetFill()->SetFillType(FillFormatType::Solid);
	cdp4->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::ForestGreen);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(cdp4);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
