
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateWaterFallChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Create a WaterFall chart to the first slide
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::WaterFall, new RectangleF(50, 50, 500, 400), false);

	//Set series text
	chart->GetChartData()->GetItem(0, 1)->SetText(L"Series 1");

	//Set category text
	std::vector<std::wstring> categories = { L"Category 1",L"Category 2", L"Category 3", L"Category 4", L"Category 5", L"Category 6", L"Category 7" };
	for (int i = 0; i < categories.size(); i++)
	{
		chart->GetChartData()->GetItem(i + 1, 0)->SetText(categories[i].c_str());
	}

	//Fill data for chart
	std::vector<double> values = { 100, 20, 50, -40, 130, -60, 70 };
	for (int i = 0; i < values.size(); i++)
	{
		chart->GetChartData()->GetItem(i + 1, 1)->SetNumberValue(values[i]);
	}

	//Set series labels
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(0, 1, 0, 1));

	//Set categories labels 
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(1, 0, categories.size(), 0));

	//Assign data to series values
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(1, 1, values.size(), 1));

	//Operate the third datapoint of first series
	intrusive_ptr<ChartDataPoint> chartDataPoint = new ChartDataPoint(chart->GetSeries()->GetItem(0));
	chartDataPoint->SetIndex(2);
	chartDataPoint->SetSetAsTotal(true);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(chartDataPoint);

	//Operate the sixth datapoint of first series
	intrusive_ptr<ChartDataPoint> chartDataPoint2 = new ChartDataPoint(chart->GetSeries()->GetItem(0));
	chartDataPoint2->SetIndex(5);
	chartDataPoint2->SetSetAsTotal(true);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(chartDataPoint2);
	chart->GetSeries()->GetItem(0)->SetShowConnectorLines(true);
	chart->GetSeries()->GetItem(0)->GetDataLabels()->SetLabelValueVisible(true);

	chart->GetChartTitle()->GetTextProperties()->SetText(L"WaterFall");
	chart->SetHasLegend(true);
	chart->GetChartLegend()->SetPosition(ChartLegendPositionType::Right);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
