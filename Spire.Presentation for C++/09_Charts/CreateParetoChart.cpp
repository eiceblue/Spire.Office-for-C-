
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateParetoChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Add line markers chart
	intrusive_ptr<RectangleF> rect1 = new RectangleF(50, 50, 500, 400);
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::Pareto, rect1, false);

	chart->GetChartData()->GetItem(0, 1)->SetText(L"series 1");

	//Set category text
	std::vector<std::wstring> categories = { L"Category 1", L"Category 2", L"Category 4", L"Category 3", L"Category 4", L"Category 2", L"Category 1", L"Category 1", L"Category 3", L"Category 2", L"Category 4", L"Category 2", L"Category 3", L"Category 1", L"Category 3", L"Category 2", L"Category 4", L"Category 1", L"Category 1", L"Category 3", L"Category 2", L"Category 4", L"Category 1", L"Category 1", L"Category 3", L"Category 2", L"Category 4", L"Category 1" };
	for (int i = 0; i < categories.size(); i++)
	{
		chart->GetChartData()->GetItem(i + 1, 0)->SetText(categories[i].c_str());
	}

	//Fill data for chart
	std::vector<double> values = { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
	for (int i = 0; i < values.size(); i++)
	{
		chart->GetChartData()->GetItem(i + 1, 1)->SetNumberValue(values[i]);
	}

	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(0, 1, 0, 1));
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(1, 0, categories.size(), 0));
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(1, 1, values.size(), 1));
	chart->GetPrimaryCategoryAxis()->SetIsBinningByCategory(true);
	chart->GetSeries()->GetItem(1)->GetLine()->GetFillFormat()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(1)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetKnownColor(KnownColors::Red);
	chart->GetChartTitle()->GetTextProperties()->SetText(L"Pareto");
	chart->SetHasLegend(true);
	chart->GetChartLegend()->SetPosition(ChartLegendPositionType::Bottom);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
