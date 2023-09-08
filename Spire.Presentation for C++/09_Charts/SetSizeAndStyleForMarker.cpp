#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SetSizeAndStyleForMarker.pptx";
	wstring outputFile = OUTPUTPATH"SetSizeAndStyleForMarker.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	for (int i = 0; i < chart->GetSeries()->GetItem(0)->GetValues()->GetCount(); i++)
	{
		//Create a ChartDataPoint object and specify the index.
		intrusive_ptr<ChartDataPoint> dataPoint = new ChartDataPoint(chart->GetSeries()->GetItem(0));
		dataPoint->SetIndex(i);

		//Set the fill color of the data marker.
		dataPoint->GetMarkerFill()->GetFill()->SetFillType(FillFormatType::Solid);
		dataPoint->GetMarkerFill()->GetFill()->GetSolidColor()->SetColor(Color::GetYellow());

		//Set the line color of the data marker.
		dataPoint->GetMarkerFill()->GetLine()->SetFillType(FillFormatType::Solid);
		dataPoint->GetMarkerFill()->GetLine()->GetSolidFillColor()->SetColor(Color::GetYellowGreen());

		//Set the size of the data marker.
		dataPoint->SetMarkerSize(20);

		//Set the style of the data marker
		dataPoint->SetMarkerStyle(ChartMarkerType::Diamond);
		chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(dataPoint);
	}

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
