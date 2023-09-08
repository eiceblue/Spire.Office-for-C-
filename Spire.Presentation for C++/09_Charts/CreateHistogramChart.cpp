
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateHistogramChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Create a Funnel chart to the first slide
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::Histogram, new RectangleF(50, 50, 500, 400), false);

	//Set series text
	chart->GetChartData()->GetItem(0, 0)->SetText(L"Series 1");

	//Fill data for chart
	std::vector<double> values = { 1, 1, 1, 3, 3, 3, 3, 5, 5, 5, 8, 8, 8, 9, 9, 9, 12, 12, 13, 13, 17, 17, 17, 19, 19, 19, 25, 25, 25, 25, 25, 25, 25, 25, 29, 29, 29, 29, 32, 32, 33, 33, 35, 35, 41, 41, 44, 45, 49, 49 };
	for (int i = 0; i < values.size(); i++)
	{
		chart->GetChartData()->GetItem(i + 1, 1)->SetNumberValue(values[i]);
	}

	//Set series label
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(0, 0, 0, 0));

	//Set values for series
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(1, 0, values.size(), 0));

	chart->GetPrimaryCategoryAxis()->SetNumberOfBins(7);
	chart->GetPrimaryCategoryAxis()->SetGapWidth(20);
	//Chart title
	chart->GetChartTitle()->GetTextProperties()->SetText(L"Histogram");
	chart->GetChartLegend()->SetPosition(ChartLegendPositionType::Bottom);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
