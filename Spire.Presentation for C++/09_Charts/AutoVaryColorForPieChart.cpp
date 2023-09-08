
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"AutoVaryColorForPieChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	intrusive_ptr<RectangleF> rect1 = new RectangleF(40, 100, 550, 320);
	//Add a pie chart
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::Pie, rect1, false);
	chart->GetChartTitle()->GetTextProperties()->SetText(L"Sales by Quarter");
	chart->GetChartTitle()->GetTextProperties()->SetIsCentered(true);
	chart->GetChartTitle()->SetHeight(30);
	chart->SetHasTitle(true);
	//Attach the data to chart
	std::vector<std::wstring> quarters = { L"1st Qtr", L"2nd Qtr", L"3rd Qtr", L"4th Qtr" };
	std::vector<int> sales = { 210, 320, 180, 500 };
	chart->GetChartData()->GetItem(0, 0)->SetText(L"Quarters");
	chart->GetChartData()->GetItem(0, 1)->SetText(L"Sales");
	for (int i = 0; i < quarters.size(); ++i)
	{
		chart->GetChartData()->GetItem(i + 1, 0)->SetText(quarters[i].c_str());
		chart->GetChartData()->GetItem(i + 1, 1)->SetNumberValue(sales[i]);
	}

	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(L"B1", L"B1"));
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(L"A2", L"A5"));
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(L"B2", L"B5"));


	//Set whether auto vary color, default value is true
	chart->GetSeries()->GetItem(0)->SetIsVaryColor(false);

	chart->GetSeries()->GetItem(0)->SetDistance(15);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
