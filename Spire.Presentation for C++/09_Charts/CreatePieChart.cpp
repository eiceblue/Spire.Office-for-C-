
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreatePieChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Add line markers chart
	intrusive_ptr<RectangleF> rect1 = new RectangleF(40, 100, 550, 320);
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::Pie, rect1, false);

	chart->GetChartTitle()->GetTextProperties()->SetText(L"Sales by Quarter");
	chart->GetChartTitle()->GetTextProperties()->SetIsCentered(true);
	chart->GetChartTitle()->SetHeight(30);
	chart->SetHasTitle(true);

	//Define some data.
	std::vector<std::wstring> quarters = { L"1st Qtr", L"2nd Qtr", L"3rd Qtr", L"4th Qtr" };
	std::vector<int> sales = { 210, 320, 180, 500 };

	//Append data to ChartData, which represents a data table where the chart data is stored.
	chart->GetChartData()->GetItem(0, 0)->SetText(L"Quarters");
	chart->GetChartData()->GetItem(0, 1)->SetText(L"Sales");
	for (int i = 0; i < quarters.size(); ++i)
	{
		chart->GetChartData()->GetItem(i + 1, 0)->SetText(quarters[i].c_str());
		chart->GetChartData()->GetItem(i + 1, 1)->SetNumberValue(sales[i]);
	}

	//Set category labels, series label and series data.
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(L"B1", L"B1"));
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(L"A2", L"A5"));
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(L"B2", L"B5"));

	//Add data points to series and fill each data point with different color.
	for (int i = 0; i < chart->GetSeries()->GetItem(0)->GetValues()->GetCount(); i++)
	{
		intrusive_ptr<ChartDataPoint> cdp = new ChartDataPoint(chart->GetSeries()->GetItem(0));
		cdp->SetIndex(i);
		chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(cdp);

	}
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(0)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::RosyBrown);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(1)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(1)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::LightBlue);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(2)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(2)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::LightPink);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(3)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(3)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::MediumPurple);

	//Set the data labels to display label value and percentage value.
	chart->GetSeries()->GetItem(0)->GetDataLabels()->SetLabelValueVisible(true);
	chart->GetSeries()->GetItem(0)->GetDataLabels()->SetPercentValueVisible(true);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
