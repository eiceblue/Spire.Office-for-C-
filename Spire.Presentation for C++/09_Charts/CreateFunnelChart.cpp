
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateFunnelChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Create a Funnel chart to the first slide
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::Funnel, new RectangleF(50, 50, 550, 400), false);

	//Set series text
	chart->GetChartData()->GetItem(0, 1)->SetText(L"Series 1");

	//Set category text
	std::vector<std::wstring> categories = { L"Website Visits", L"Download", L"Uploads", L"Requested price", L"Invoice sent", L"Finalized" };
	for (int i = 0; i < categories.size(); i++)
	{
		chart->GetChartData()->GetItem(i + 1, 0)->SetText(categories[i].c_str());
	}

	//Fill data for chart
	std::vector<double> values = { 50000, 47000, 30000, 15000, 9000, 5600 };
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

	//Set the chart title
	chart->GetChartTitle()->GetTextProperties()->SetText(L"Funnel");

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
