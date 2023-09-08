
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateLineMarkersChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Add line markers chart
	intrusive_ptr<RectangleF> rect1 = new RectangleF(90, 100, 550, 320);
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::LineMarkers, rect1, false);

	//Chart title
	chart->GetChartTitle()->GetTextProperties()->SetText(L"Line Makers Chart");
	chart->GetChartTitle()->GetTextProperties()->SetIsCentered(true);
	chart->GetChartTitle()->SetHeight(30);
	chart->SetHasTitle(true);

	//Data for series
	std::vector<double> Series1 = { 7.7, 8.9, 1.0, 2.4 };
	std::vector<double> Series2 = { 15.2, 5.3, 6.7, 8 };

	//Set series text
	chart->GetChartData()->GetItem(0, 1)->SetText(L"Series1");
	chart->GetChartData()->GetItem(0, 2)->SetText(L"Series2");

	//Set category text
	chart->GetChartData()->GetItem(1, 0)->SetText(L"Category 1");
	chart->GetChartData()->GetItem(2, 0)->SetText(L"Category 2");
	chart->GetChartData()->GetItem(3, 0)->SetText(L"Category 3");
	chart->GetChartData()->GetItem(4, 0)->SetText(L"Category 4");

	//Fill data for chart
	for (int i = 0; i < Series1.size(); ++i)
	{
		chart->GetChartData()->GetItem(i + 1, 1)->SetNumberValue(Series1[i]);
		chart->GetChartData()->GetItem(i + 1, 2)->SetNumberValue(Series2[i]);

	}
	//Set series label
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(L"B1", L"C1"));
	//Set category label
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(L"A2", L"A5"));

	//Set values for series
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(L"B2", L"B5"));
	chart->GetSeries()->GetItem(1)->SetValues(chart->GetChartData()->GetItem(L"C2", L"C5"));

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
