
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateSunBurstChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Create a SunBurst chart to the first slide
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::SunBurst, new RectangleF(50, 50, 500, 400), false);

	//Set series text
	chart->GetChartData()->GetItem(0, 3)->SetText(L"Series 1");

	//Set category text
	std::vector<std::vector<std::wstring>> categories =
	{
		{L"Branch 1", L"Stem 1", L"Leaf 1"},
		{L"Branch 1", L"Stem 1", L"Leaf 2"},
		{L"Branch 1", L"Stem 1", L"Leaf 3"},
		{L"Branch 1", L"Stem 2", L"Leaf 4"},
		{L"Branch 1", L"Stem 2", L"Leaf 5"},
		{L"Branch 1", L"Leaf 6", L""},
		{L"Branch 1", L"Leaf 7", L""},
		{L"Branch 2", L"Stem 3", L"Leaf 8"},
		{L"Branch 2", L"Leaf 9", L""},
		{L"Branch 2", L"Stem 4", L"Leaf 10"},
		{L"Branch 2", L"Stem 4", L"Leaf 11"},
		{L"Branch 2", L"Stem 5", L"Leaf 12"},
		{L"Branch 3", L"Stem 5", L"Leaf 13"},
		{L"Branch 3", L"Stem 6", L"Leaf 14"},
		{L"Branch 3", L"Leaf 15", L""}
	};
	for (int i = 0; i < 15; i++)
	{
		for (int j = 0; j < 3; j++)
		{
			chart->GetChartData()->GetItem(i + 1, j)->SetText(categories[i][j].c_str());
		}
	}
	//Fill data for chart
	std::vector<double> values = { 17, 23, 48, 22, 76, 54, 77, 26, 44, 63, 10, 15, 48, 15, 51 };
	for (int i = 0; i < values.size(); i++)
	{
		chart->GetChartData()->GetItem(i + 1, 3)->SetNumberValue(values[i]);
	}

	//Set series labels
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(0, 3, 0, 3));

	//Set categories labels 
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(1, 0, values.size(), 2));

	//Assign data to series values
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(1, 3, values.size(), 3));

	chart->GetSeries()->GetItem(0)->GetDataLabels()->SetCategoryNameVisible(true);
	chart->GetChartTitle()->GetTextProperties()->SetText(L"SunBurst");
	chart->SetHasLegend(true);
	chart->GetChartLegend()->SetPosition(ChartLegendPositionType::Top);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
