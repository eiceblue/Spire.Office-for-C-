
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_2.pptx";
	wstring outputFile = OUTPUTPATH"SetNumberFormatAndRemoveTickMarksOfChart.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Set percentage number format for the axis value of chart.
	chart->GetPrimaryValueAxis()->SetNumberFormat(L"0#\\%");

	//Remove the tick marks for value axis and category axis.
	chart->GetPrimaryValueAxis()->SetMajorTickMark(TickMarkType::TickMarkNone);
	chart->GetPrimaryValueAxis()->SetMinorTickMark(TickMarkType::TickMarkNone);
	chart->GetPrimaryCategoryAxis()->SetMajorTickMark(TickMarkType::TickMarkNone);
	chart->GetPrimaryCategoryAxis()->SetMinorTickMark(TickMarkType::TickMarkNone);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
