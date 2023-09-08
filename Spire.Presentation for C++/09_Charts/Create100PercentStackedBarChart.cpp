
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"Create100PercentStackedBarChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Add a "Bar100PercentStacked" chart to the first slide.
	ppt->GetSlideSize()->SetType(SlideSizeType::Screen16x9);
	intrusive_ptr<SizeF> slidesize = ppt->GetSlideSize()->GetSize();

	auto slide = ppt->GetSlides()->GetItem(0);

	//Append a chart.
	intrusive_ptr<RectangleF> rect = new RectangleF(20, 20, slidesize->GetWidth() - 40, slidesize->GetHeight() - 40);
	intrusive_ptr<IChart> chart = slide->GetShapes()->AppendChart(ChartType::Bar100PercentStacked, rect);

	//Write data to the chart data.
	std::vector<std::wstring> columnlabels = { L"Series 1", L"Series 2", L"Series 3" };

	//Insert the column labels.
	for (int c = 0; c < columnlabels.size(); ++c)
	{
		chart->GetChartData()->GetItem(0, c + 1)->SetText(columnlabels[c].c_str());
	}

	std::vector<std::wstring> rowlabels = { L"Category 1", L"Category 2", L"Category 3" };

	//Insert the row labels.
	for (int r = 0; r < rowlabels.size(); ++r)
	{
		chart->GetChartData()->GetItem(r + 1, 0)->SetText(rowlabels[r].c_str());
	}

	std::vector<std::vector<double>> values =
	{
		{20.83233, 10.34323, -10.354667},
		{10.23456, -12.23456, 23.34456},
		{12.34345, -23.34343, -13.23232}
	};

	//Insert the values.
	double value = 0.0;
	for (int r = 0; r < rowlabels.size(); ++r)
	{
		for (int c = 0; c < columnlabels.size(); ++c)
		{
			int temp = values[r][c] > 0.0 ? (values[r][c] + 0.005) * 100 : (values[r][c] - 0.005) * 100;
			value = temp / 100.0;
			chart->GetChartData()->GetItem(r + 1, c + 1)->SetNumberValue(value);
		}
	}

	chart->GetSeries()->SetSeriesLabel(chart->GetChartData()->GetItem(0, 1, 0, columnlabels.size()));
	chart->GetCategories()->SetCategoryLabels(chart->GetChartData()->GetItem(1, 0, rowlabels.size(), 0));

	//Set the position of category axis.
	chart->GetPrimaryCategoryAxis()->SetPosition(AxisPositionType::Left);
	chart->GetSecondaryCategoryAxis()->SetPosition(AxisPositionType::Left);
	chart->GetPrimaryCategoryAxis()->SetTickLabelPosition(TickLabelPositionType::TickLabelPositionLow);

	//Set the data, font and format for the series of each column.
	for (int c = 0; c < columnlabels.size(); ++c)
	{
		chart->GetSeries()->GetItem(c)->SetValues(chart->GetChartData()->GetItem(1, c + 1, rowlabels.size(), c + 1));
		chart->GetSeries()->GetItem(c)->GetFill()->SetFillType(FillFormatType::Solid);
		chart->GetSeries()->GetItem(c)->SetInvertIfNegative(false);

		for (int r = 0; r < rowlabels.size(); ++r)
		{
			auto label = chart->GetSeries()->GetItem(c)->GetDataLabels()->Add();
			label->SetLabelValueVisible(true);
			chart->GetSeries()->GetItem(c)->GetDataLabels()->GetItem(r)->SetHasDataSource(false);
			chart->GetSeries()->GetItem(c)->GetDataLabels()->GetItem(r)->SetNumberFormat(L"0#\\%");
			chart->GetSeries()->GetItem(c)->GetDataLabels()->GetTextProperties()->GetParagraphs()->GetItem(0)->GetDefaultCharacterProperties()->SetFontHeight(12);
		}
	}

	//Set the color of the Series.
	chart->GetSeries()->GetItem(0)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::YellowGreen);
	chart->GetSeries()->GetItem(1)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Red);
	chart->GetSeries()->GetItem(2)->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Green);

	intrusive_ptr<TextFont> font = new TextFont(L"Tw Cen MT");

	//Set the font and size for chartlegend.
	for (int k = 0; k < chart->GetChartLegend()->GetEntryTextProperties().size(); k++)
	{
		chart->GetChartLegend()->GetEntryTextProperties()[k]->SetLatinFont(font);
		chart->GetChartLegend()->GetEntryTextProperties()[k]->SetFontHeight(20);
	}

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
