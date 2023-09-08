#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetChartDataLabelRange.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Add a ColumnStacked chart
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->
		AppendChart(ChartType::ColumnStacked, new RectangleF(100, 100, 500, 400));

	//Set data for the chart
	intrusive_ptr<CellRange> cellRange = chart->GetChartData()->GetItem(L"F1");
	cellRange->SetText(L"labelA");
	cellRange = chart->GetChartData()->GetItem(L"F2");
	cellRange->SetText(L"labelB");
	cellRange = chart->GetChartData()->GetItem(L"F3");
	cellRange->SetText(L"labelC");
	cellRange = chart->GetChartData()->GetItem(L"F4");
	cellRange->SetText(L"labelD");

	//Set data label ranges
	chart->GetSeries()->GetItem(0)->SetDataLabelRanges(chart->GetChartData()->GetItem(L"F1", L"F4"));

	//Add data label
	intrusive_ptr<ChartDataLabel> dataLabel1 = chart->GetSeries()->GetItem(0)->GetDataLabels()->Add();
	dataLabel1->SetID(0);
	//Show the value
	dataLabel1->SetLabelValueVisible(false);
	//Show the label string
	dataLabel1->SetShowDataLabelsRange(true);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
