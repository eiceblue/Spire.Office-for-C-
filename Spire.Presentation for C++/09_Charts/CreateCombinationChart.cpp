
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CombinationChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	std::wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect2 = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect2);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetKnownColor(KnownColors::FloralWhite);

	//Insert a column clustered chart
	intrusive_ptr<RectangleF> rect = new RectangleF(100, 100, 550, 320);
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::ColumnClustered, rect);

	//Set chart title
	chart->GetChartTitle()->GetTextProperties()->SetText(L"Monthly Sales Report");
	chart->GetChartTitle()->GetTextProperties()->SetIsCentered(true);
	chart->GetChartTitle()->SetHeight(30);
	chart->SetHasTitle(true);

	//data
	std::vector<std::wstring> series = { L"Month",L"Sales",L"Growth rate" };
	std::vector<std::wstring> months = { L"January",L"February",L"March",L"April",L"May",L"June" };
	std::vector<int> sales = { 200,250,300,150,200,400 };
	std::vector<double> rates = { 0.6,0.8,0.6,0.2,0.5,0.9 };
	//series
	for (int c = 0; c < series.size(); c++)
	{
		chart->GetChartData()->GetItem(0, c)->SetText(series[c].c_str());
	}
	//Month
	for (int c = 0; c < months.size(); c++)
	{
		chart->GetChartData()->GetItem(c + 1, 0)->SetText(months[c].c_str());
	}
	//Sales
	for (int c = 0; c < sales.size(); c++)
	{
		chart->GetChartData()->GetItem(c + 1, 1)->SetNumberValue(sales[c]);
	}
	//rates
	for (int c = 0; c < rates.size(); c++)
	{
		chart->GetChartData()->GetItem(c + 1, 2)->SetNumberValue(rates[c]);
	}

	//Set series labels
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(L"B1", L"C1"));

	//Set categories labels    
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(L"A2", L"A7"));

	//Assign data to series values
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(L"B2", L"B7"));
	chart->GetSeries()->GetItem(1)->SetValues(chart->GetChartData()->GetItem(L"C2", L"C7"));

	//Change the chart type of serie 2 to line with markers
	chart->GetSeries()->GetItem(1)->SetType(ChartType::LineMarkers);

	//Plot data of series 2 on the secondary axis
	chart->GetSeries()->GetItem(1)->SetUseSecondAxis(true);

	//Set the number format as percentage 
	chart->GetSecondaryValueAxis()->SetNumberFormat(L"0%");

	//Hide gridlinkes of secondary axis
	chart->GetSecondaryValueAxis()->GetMajorGridTextLines()->SetFillType(FillFormatType::None);

	//Set overlap
	chart->SetOverLap(-50);

	//Set gapwidth
	chart->SetGapWidth(200);
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
