
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"MultipleCategoryChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Add line markers chart
	intrusive_ptr<RectangleF> rect1 = new RectangleF(90, 100, 550, 320);
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::ColumnClustered, rect1, false);

	//Chart title
	chart->GetChartTitle()->GetTextProperties()->SetText(L"Muli-Category");
	chart->GetChartTitle()->GetTextProperties()->SetIsCentered(true);
	chart->GetChartTitle()->SetHeight(30);
	chart->SetHasTitle(true);


	//Data for series
	std::vector<double> Series1 = { 7.7, 8.9, 7, 6, 7, 8 };

	//Set series text
	chart->GetChartData()->GetItem(0, 2)->SetText(L"Series1");

	//Set category text
	chart->GetChartData()->GetItem(1, 0)->SetText(L"Grp 1");
	chart->GetChartData()->GetItem(3, 0)->SetText(L"Grp 2");
	chart->GetChartData()->GetItem(5, 0)->SetText(L"Grp 3");

	chart->GetChartData()->GetItem(1, 1)->SetText(L"A");
	chart->GetChartData()->GetItem(2, 1)->SetText(L"B");
	chart->GetChartData()->GetItem(3, 1)->SetText(L"C");
	chart->GetChartData()->GetItem(4, 1)->SetText(L"D");
	chart->GetChartData()->GetItem(5, 1)->SetText(L"E");
	chart->GetChartData()->GetItem(6, 1)->SetText(L"F");


	//Fill data for chart
	for (int i = 0; i < Series1.size(); ++i)
	{
		chart->GetChartData()->GetItem(i + 1, 2)->SetNumberValue(Series1[i]);

	}

	//Set series label
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(L"C1", L"C1"));
	//Set category label
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(L"A2", L"B7"));

	//Set values for series
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(L"C2", L"C7"));

	//Set if the category axis has multiple levels
	chart->GetPrimaryCategoryAxis()->SetHasMultiLvlLbl(true);
	//Merge same label
	chart->GetPrimaryCategoryAxis()->SetIsMergeSameLabel(true);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
