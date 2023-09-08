
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateScatterChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	std::wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect2 = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect2);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetKnownColor(KnownColors::FloralWhite);

	//Insert a chart and set chart title and chart type
	intrusive_ptr<RectangleF> rect1 = new RectangleF(90, 100, 550, 320);
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::ScatterMarkers, rect1, false);
	chart->GetChartTitle()->GetTextProperties()->SetText(L"ScatterMarker Chart");
	chart->GetChartTitle()->GetTextProperties()->SetIsCentered(true);
	chart->GetChartTitle()->SetHeight(30);
	chart->SetHasTitle(true);

	//Set chart data
	std::vector<double> xdata = { 2.7, 8.9, 10.0, 12.4 };
	std::vector<double> ydata = { 3.2, 15.3, 6.7, 8 };

	chart->GetChartData()->GetItem(0, 0)->SetText(L"X-Value");
	chart->GetChartData()->GetItem(0, 1)->SetText(L"Y-Value");

	for (int i = 0; i < xdata.size(); ++i)
	{
		chart->GetChartData()->GetItem(i + 1, 0)->SetNumberValue(xdata[i]);
		chart->GetChartData()->GetItem(i + 1, 1)->SetNumberValue(ydata[i]);
	}

	//Set the series label
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(L"B1", L"B1"));

	//Assign data to X axis, Y axis and Bubbles
	chart->GetSeries()->GetItem(0)->SetXValues(chart->GetChartData()->GetItem(L"A2", L"A5"));
	chart->GetSeries()->GetItem(0)->SetYValues(chart->GetChartData()->GetItem(L"B2", L"B5"));

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
