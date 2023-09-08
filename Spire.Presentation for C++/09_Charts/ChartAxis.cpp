
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChartAxis.pptx";
	wstring outputFile = OUTPUTPATH"ChartAxis.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get chart on the first slide
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Add a secondary axis to display the value of Series 3
	chart->GetSeries()->GetItem(2)->SetUseSecondAxis(true);

	//Set the grid line of secondary axis as invisible
	chart->GetSecondaryValueAxis()->GetMajorGridTextLines()->SetFillType(FillFormatType::None);

	//Set bounds of axis value. Before we assign values, we must set IsAutoMax and IsAutoMin as false, otherwise MS PowerPoint will automatically set the values.
	chart->GetPrimaryValueAxis()->SetIsAutoMax(false);
	chart->GetPrimaryValueAxis()->SetIsAutoMin(false);
	chart->GetSecondaryValueAxis()->SetIsAutoMax(false);
	chart->GetSecondaryValueAxis()->SetIsAutoMax(false);

	chart->GetPrimaryValueAxis()->SetMinValue(0.0);
	chart->GetPrimaryValueAxis()->SetMaxValue(5.0);
	chart->GetSecondaryValueAxis()->SetMinValue(0.0);
	chart->GetSecondaryValueAxis()->SetMaxValue(1.0);

	//Set axis line format
	chart->GetPrimaryValueAxis()->GetMinorGridLines()->SetFillType(FillFormatType::Solid);
	chart->GetSecondaryValueAxis()->GetMinorGridLines()->SetFillType(FillFormatType::Solid);
	chart->GetPrimaryValueAxis()->GetMinorGridLines()->SetWidth(0.1);
	chart->GetSecondaryValueAxis()->GetMinorGridLines()->SetWidth(0.1);
	chart->GetPrimaryValueAxis()->GetMinorGridLines()->GetSolidFillColor()->SetColor(Color::GetLightGray());
	chart->GetSecondaryValueAxis()->GetMinorGridLines()->GetSolidFillColor()->SetColor(Color::GetLightGray());
	chart->GetPrimaryValueAxis()->GetMinorGridLines()->SetDashStyle(LineDashStyleType::Dash);
	chart->GetSecondaryValueAxis()->GetMinorGridLines()->SetDashStyle(LineDashStyleType::Dash);

	chart->GetPrimaryValueAxis()->GetMajorGridTextLines()->SetWidth(0.3);
	chart->GetPrimaryValueAxis()->GetMajorGridTextLines()->GetSolidFillColor()->SetColor(Color::GetLightSkyBlue());
	chart->GetSecondaryValueAxis()->GetMajorGridTextLines()->SetWidth(0.3);
	chart->GetSecondaryValueAxis()->GetMajorGridTextLines()->GetSolidFillColor()->SetColor(Color::GetLightSkyBlue());

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
