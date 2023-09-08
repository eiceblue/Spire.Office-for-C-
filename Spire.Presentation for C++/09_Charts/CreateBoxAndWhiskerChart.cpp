
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateBoxAndWhiskerChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	// Insert a BoxAndWhisker chart to the first slide 
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::BoxAndWhisker, new RectangleF(50, 50, 500, 400), false);

	// Series labels
	std::vector<std::wstring> seriesLabel = { L"Series 1", L"Series 2", L"Series 3" };
	for (int i = 0; i < seriesLabel.size(); i++)
	{
		chart->GetChartData()->GetItem(0, i + 1)->SetText(L"Series 1");
	}

	// Categories
	std::vector<std::wstring> categories = { L"Category 1", L"Category 1", L"Category 1", L"Category 1",
		L"Category 1", L"Category 1", L"Category 1", L"Category 2", L"Category 2", L"Category 2",
		L"Category 2", L"Category 2", L"Category 2", L"Category 3", L"Category 3", L"Category 3",
		L"Category 3", L"Category 3" };
	for (int i = 0; i < categories.size(); i++)
	{
		chart->GetChartData()->GetItem(i + 1, 0)->SetText(categories[i].c_str());
	}
	// Values
	std::vector<std::vector<double>> values =
	{
		{-7, -3, -24},
		{-10, 1, 11},
		{-28, -6, 34},
		{47, 2, -21},
		{35, 17, 22},
		{-22, 15, 19},
		{17, -11, 25},
		{-30, 18, 25},
		{49, 22, 56},
		{37, 22, 15},
		{-55, 25, 31},
		{14, 18, 22},
		{18, -22, 36},
		{-45, 25, -17},
		{-33, 18, 22},
		{18, 2, -23},
		{-33, -22, 10},
		{10, 19, 22}
	};
	for (int i = 0; i < seriesLabel.size(); i++)
	{
		for (int j = 0; j < categories.size(); j++)
		{
			chart->GetChartData()->GetItem(j + 1, i + 1)->SetNumberValue(values[j][i]);
		}
	}
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(0, 1, 0, seriesLabel.size()));
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(1, 0, categories.size(), 0));

	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(1, 1, categories.size(), 1));
	chart->GetSeries()->GetItem(1)->SetValues(chart->GetChartData()->GetItem(1, 2, categories.size(), 2));
	chart->GetSeries()->GetItem(2)->SetValues(chart->GetChartData()->GetItem(1, 3, categories.size(), 3));

	chart->GetSeries()->GetItem(0)->SetShowInnerPoints(false);
	chart->GetSeries()->GetItem(0)->SetShowOutlierPoints(true);
	chart->GetSeries()->GetItem(0)->SetShowMeanMarkers(true);
	chart->GetSeries()->GetItem(0)->SetShowMeanLine(true);
	chart->GetSeries()->GetItem(0)->SetQuartileCalculationType(QuartileCalculation::ExclusiveMedian);

	chart->GetSeries()->GetItem(1)->SetShowInnerPoints(false);
	chart->GetSeries()->GetItem(1)->SetShowOutlierPoints(true);
	chart->GetSeries()->GetItem(1)->SetShowMeanMarkers(true);
	chart->GetSeries()->GetItem(1)->SetShowMeanLine(true);
	chart->GetSeries()->GetItem(1)->SetQuartileCalculationType(QuartileCalculation::InclusiveMedian);

	chart->GetSeries()->GetItem(2)->SetShowInnerPoints(false);
	chart->GetSeries()->GetItem(2)->SetShowOutlierPoints(true);
	chart->GetSeries()->GetItem(2)->SetShowMeanMarkers(true);
	chart->GetSeries()->GetItem(2)->SetShowMeanLine(true);
	chart->GetSeries()->GetItem(2)->SetQuartileCalculationType(QuartileCalculation::ExclusiveMedian);

	chart->SetHasLegend(true);
	chart->GetChartTitle()->GetTextProperties()->SetText(L"BoxAndWhisker");
	chart->GetChartLegend()->SetPosition(ChartLegendPositionType::Top);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
