
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"BubbleChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	std::wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect2 = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect2);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetKnownColor(KnownColors::FloralWhite);

	//Add bubble chart
	intrusive_ptr<RectangleF> rect1 = new RectangleF(90, 100, 550, 320);
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::Bubble, rect1, false);

	//Chart title
	chart->GetChartTitle()->GetTextProperties()->SetText(L"Bubble Chart");
	chart->GetChartTitle()->GetTextProperties()->SetIsCentered(true);
	chart->GetChartTitle()->SetHeight(30);
	chart->SetHasTitle(true);

	//Attach the data to chart
	std::vector<double> xdata = { 7.7, 8.9, 1.0, 2.4 };
	std::vector<double> ydata = { 15.2, 5.3, 6.7, 8 };
	std::vector<double> size = { 1.1, 2.4, 3.7, 4.8 };

	chart->GetChartData()->GetItem(0, 0)->SetText(L"X-Value");
	chart->GetChartData()->GetItem(0, 1)->SetText(L"Y-Value");
	chart->GetChartData()->GetItem(0, 2)->SetText(L"Size");

	for (int i = 0; i < xdata.size(); ++i)
	{
		chart->GetChartData()->GetItem(i + 1, 0)->SetNumberValue(xdata[i]);
		chart->GetChartData()->GetItem(i + 1, 1)->SetNumberValue(ydata[i]);
		chart->GetChartData()->GetItem(i + 1, 2)->SetNumberValue(size[i]);
	}

	//Set series label
	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(L"B1", L"B1"));

	chart->GetSeries()->GetItem(0)->SetXValues(chart->GetChartData()->GetItem(L"A2", L"A5"));
	chart->GetSeries()->GetItem(0)->SetYValues(chart->GetChartData()->GetItem(L"B2", L"B5"));
	chart->GetSeries()->GetItem(0)->GetBubbles()->Add(chart->GetChartData()->GetItem(L"C2"));
	chart->GetSeries()->GetItem(0)->GetBubbles()->Add(chart->GetChartData()->GetItem(L"C3"));
	chart->GetSeries()->GetItem(0)->GetBubbles()->Add(chart->GetChartData()->GetItem(L"C4"));
	chart->GetSeries()->GetItem(0)->GetBubbles()->Add(chart->GetChartData()->GetItem(L"C5"));

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
