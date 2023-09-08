#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"VaryColorsOfSameSeriesDataMarkers.pptx";
	wstring outputFile = OUTPUTPATH"VaryColorOfSameSerieDataMarkers.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Create a ChartDataPoint object and specify the index.
	intrusive_ptr<ChartDataPoint> dataPoint = new ChartDataPoint(chart->GetSeries()->GetItem(0));
	dataPoint->SetIndex(0);

	//Set the fill color of the data marker.
	dataPoint->GetMarkerFill()->GetFill()->SetFillType(FillFormatType::Solid);
	dataPoint->GetMarkerFill()->GetFill()->GetSolidColor()->SetColor(Color::GetRed());

	//Set the line color of the data marker.
	dataPoint->GetMarkerFill()->GetLine()->SetFillType(FillFormatType::Solid);
	dataPoint->GetMarkerFill()->GetLine()->GetSolidFillColor()->SetColor(Color::GetRed());

	//Add the data point to the point collection of a series.
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(dataPoint);

	dataPoint = new ChartDataPoint(chart->GetSeries()->GetItem(0));
	dataPoint->SetIndex(1);
	dataPoint->GetMarkerFill()->GetFill()->SetFillType(FillFormatType::Solid);
	dataPoint->GetMarkerFill()->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());
	dataPoint->GetMarkerFill()->GetLine()->SetFillType(FillFormatType::Solid);
	dataPoint->GetMarkerFill()->GetLine()->GetSolidFillColor()->SetColor(Color::GetBlack());
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(dataPoint);

	dataPoint = new ChartDataPoint(chart->GetSeries()->GetItem(0));
	dataPoint->SetIndex(2);
	dataPoint->GetMarkerFill()->GetFill()->SetFillType(FillFormatType::Solid);
	dataPoint->GetMarkerFill()->GetFill()->GetSolidColor()->SetColor(Color::GetBlue());
	dataPoint->GetMarkerFill()->GetLine()->SetFillType(FillFormatType::Solid);
	dataPoint->GetMarkerFill()->GetLine()->GetSolidFillColor()->SetColor(Color::GetBlue());
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(dataPoint);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
