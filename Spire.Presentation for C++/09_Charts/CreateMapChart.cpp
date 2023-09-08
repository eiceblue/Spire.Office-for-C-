
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateMapChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Add line markers chart
	intrusive_ptr<RectangleF> rect1 = new RectangleF(50, 50, 450, 450);
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::Map, rect1, false);

	chart->GetChartData()->GetItem(0, 1)->SetText(L"series");

	//Define some data.
	std::vector<std::wstring> countries = { L"China", L"Russia", L"France", L"Mexico", L"United States", L"India", L"Australia" };
	for (int i = 0; i < countries.size(); i++)
	{
		chart->GetChartData()->GetItem(i + 1, 0)->SetText(countries[i].c_str());
	}
	std::vector<int> values = { 32, 20, 23, 17, 18, 6, 11 };
	for (int i = 0; i < values.size(); i++)
	{
		chart->GetChartData()->GetItem(i + 1, 1)->SetNumberValue(values[i]);
	}
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(0, 1, 0, 1));
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(1, 0, 7, 0));
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(1, 1, 7, 1));

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
