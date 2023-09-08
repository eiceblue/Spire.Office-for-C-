#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"DoughnutChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	intrusive_ptr<RectangleF> rect = new RectangleF(80, 100, 550, 320);

	//Set background image
	std::wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect2 = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect2);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetKnownColor(KnownColors::FloralWhite);
	//Add a Doughnut chart
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::Doughnut, rect, false);
	chart->GetChartTitle()->GetTextProperties()->SetText(L"Market share by country");
	chart->GetChartTitle()->GetTextProperties()->SetIsCentered(true);
	chart->GetChartTitle()->SetHeight(30);

	std::vector<std::wstring> countries = { L"Guba", L"Mexico", L"France", L"German" };
	std::vector<int> sales = { 1800, 3000, 5100, 6200 };
	chart->GetChartData()->GetItem(0, 0)->SetText(L"Countries");
	chart->GetChartData()->GetItem(0, 1)->SetText(L"Sales");
	for (int i = 0; i < countries.size(); ++i)
	{
		chart->GetChartData()->GetItem(i + 1, 0)->SetText(countries[i].c_str());
		chart->GetChartData()->GetItem(i + 1, 1)->SetNumberValue(sales[i]);
	}
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(L"B1", L"B1"));
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(L"A2", L"A5"));
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(L"B2", L"B5"));

	for (int i = 0; i < chart->GetSeries()->GetItem(0)->GetValues()->GetCount(); i++)
	{
		intrusive_ptr<ChartDataPoint> cdp = new ChartDataPoint(chart->GetSeries()->GetItem(0));
		cdp->SetIndex(i);
		chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(cdp);
	}
	//Set the series color
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(0)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::LightBlue);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(1)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(1)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::MediumPurple);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(2)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(2)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::DarkGray);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(3)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(0)->GetDataPoints()->GetItem(3)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::DarkOrange);

	chart->GetSeries()->GetItem(0)->GetDataLabels()->SetLabelValueVisible(true);
	chart->GetSeries()->GetItem(0)->GetDataLabels()->SetPercentValueVisible(true);
	chart->GetSeries()->GetItem(0)->SetDoughnutHoleSize(60);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
