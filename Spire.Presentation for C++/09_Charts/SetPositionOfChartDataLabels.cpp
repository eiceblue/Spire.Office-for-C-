#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_2.pptx";
	wstring outputFile = OUTPUTPATH"SetPositionOfChartDataLabels.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Add data label to chart and set its id.
	intrusive_ptr<ChartDataLabel> label1 = chart->GetSeries()->GetItem(0)->GetDataLabels()->Add();
	label1->SetID(0);

	//Set the default position of data label. This position is relative to the data markers.
	//label1.Position = ChartDataLabelPosition.OutsideEnd;

	//Set custom position of data label. This position is relative to the default position.
	label1->SetX(0.1);
	label1->SetY(-0.1);

	//Set label value visible
	label1->SetLabelValueVisible(true);

	//Set legend key invisible
	label1->SetLegendKeyVisible(false);

	//Set category name invisible
	label1->SetCategoryNameVisible(false);

	//Set series name invisible
	label1->SetSeriesNameVisible(false);

	//Set Percentage invisible
	label1->SetPercentageVisible(false);

	//Set border style and fill style of data label
	label1->GetLine()->SetFillType(FillFormatType::Solid);
	label1->GetLine()->GetSolidFillColor()->SetColor(Color::GetBlue());
	label1->GetFill()->SetFillType(FillFormatType::Solid);
	label1->GetFill()->GetSolidColor()->SetColor(Color::GetOrange());

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
