
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CreateCylinder3DClusteredChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	std::wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect2 = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect2);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetKnownColor(KnownColors::FloralWhite);

	//Insert chart
	intrusive_ptr<RectangleF> rect = new RectangleF(ppt->GetSlideSize()->GetSize()->GetWidth() / 2 - 200, 85, 400, 400);
	intrusive_ptr<IChart> chart = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendChart(ChartType::Cylinder3DClustered, rect);

	//Add chart Title
	chart->GetChartTitle()->GetTextProperties()->SetText(L"Report");
	chart->GetChartTitle()->GetTextProperties()->SetIsCentered(true);
	chart->GetChartTitle()->SetHeight(30);
	chart->SetHasTitle(true);

	//data
	std::vector<std::wstring> series = { L"SalesPers",L"SaleAmt",L"ComPct",L"ComAmt" };
	std::vector<std::wstring> SalesPers = { L"Joe",L"Robert",L"Michelle",L"Erich",L"Dafna",L"Rob",L"Total" };
	std::vector<int> SaleAmt = { 260,660,940,410,800,900,3970 };
	std::vector<double> ComPct = { 10,15,15, 12,15,15,14.000000000000002 , };
	std::vector<double> ComAmt = { 26,99,141,49.2, 120,135,95.03 };
	//series
	for (int c = 0; c < series.size(); c++)
	{
		chart->GetChartData()->GetItem(0, c)->SetText(series[c].c_str());
	}
	//SalesPers
	for (int c = 0; c < SalesPers.size(); c++)
	{
		chart->GetChartData()->GetItem(c + 1, 0)->SetText(SalesPers[c].c_str());
	}
	//SaleAmt
	for (int c = 0; c < SaleAmt.size(); c++)
	{
		chart->GetChartData()->GetItem(c + 1, 1)->SetNumberValue(SaleAmt[c]);
	}
	//ComPct
	for (int c = 0; c < ComPct.size(); c++)
	{
		chart->GetChartData()->GetItem(c + 1, 2)->SetNumberValue(ComPct[c]);
	}
	//ComAmt
	for (int c = 0; c < ComAmt.size(); c++)
	{
		chart->GetChartData()->GetItem(c + 1, 2)->SetNumberValue(ComAmt[c]);
	}

	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(L"B1", L"D1"));
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(L"A2", L"A7"));
	chart->GetSeries()->GetItem(0)->SetValues(chart->GetChartData()->GetItem(L"B2", L"B7"));
	chart->GetSeries()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(0)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Brown);
	chart->GetSeries()->GetItem(1)->SetValues(chart->GetChartData()->GetItem(L"C2", L"C7"));
	chart->GetSeries()->GetItem(1)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(1)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Green);
	chart->GetSeries()->GetItem(2)->SetValues(chart->GetChartData()->GetItem(L"D2", L"D7"));
	chart->GetSeries()->GetItem(2)->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetSeries()->GetItem(2)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Orange);

	//Set the 3D rotation
	chart->GetRotationThreeD()->SetXDegree(10);
	chart->GetRotationThreeD()->SetYDegree(10);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
