
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddAndFormatErrorBars.pptx";
	wstring outputFile = OUTPUTPATH"AddAndFormatErrorBars.pptx";

	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the column chart on the first slide and set chart title.
	intrusive_ptr<IChart> columnChart = Object::Dynamic_cast<IChart>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	columnChart->GetChartTitle()->GetTextProperties()->SetText(L"Vertical Error Bars");

	//Add Y (Vertical) Error Bars.

	//Get Y error bars of the first chart series.
	intrusive_ptr<IErrorBarsFormat> errorBarsYFormat1 = columnChart->GetSeries()->GetItem(0)->GetErrorBarsYFormat();

	//Set end cap.
	errorBarsYFormat1->SetErrorBarNoEndCap(false);

	//Specify direction.
	errorBarsYFormat1->SetErrorBarSimType(ErrorBarSimpleType::Plus);

	//Specify error amount type.
	errorBarsYFormat1->SetErrorBarvType(ErrorValueType::StandardError);

	//Set value.
	errorBarsYFormat1->SetErrorBarVal(0.3);

	//Set line format.
	errorBarsYFormat1->GetLine()->SetFillType(FillFormatType::Solid);
	errorBarsYFormat1->GetLine()->GetSolidFillColor()->SetKnownColor(KnownColors::MediumVioletRed);
	errorBarsYFormat1->GetLine()->SetWidth(1);

	//Get the bubble chart on the second slide and set chart title.
	intrusive_ptr<IChart> bubbleChart = Object::Dynamic_cast<IChart>(presentation->GetSlides()
		->GetItem(1)->GetShapes()->GetItem(0));

	bubbleChart->GetChartTitle()->GetTextProperties()->SetText(L"Vertical and Horizontal Error Bars");


	//Add X (Horizontal) and Y (Vertical) Error Bars.
	//Get X error bars of the first chart series.
	intrusive_ptr<IErrorBarsFormat> errorBarsXFormat = bubbleChart->GetSeries()->GetItem(0)->GetErrorBarsXFormat();

	//Set end cap.
	errorBarsXFormat->SetErrorBarNoEndCap(false);

	//Specify direction.
	errorBarsXFormat->SetErrorBarSimType(ErrorBarSimpleType::Both);

	//Specify error amount type.
	errorBarsXFormat->SetErrorBarvType(ErrorValueType::StandardError);

	//Set value.
	errorBarsXFormat->SetErrorBarVal(0.3);

	//Get Y error bars of the first chart series.
	intrusive_ptr<IErrorBarsFormat> errorBarsYFormat2 = bubbleChart->GetSeries()->GetItem(0)->GetErrorBarsYFormat();

	//Set end cap.
	errorBarsYFormat2->SetErrorBarNoEndCap(false);

	//Specify direction.
	errorBarsYFormat2->SetErrorBarSimType(ErrorBarSimpleType::Both);

	//Specify error amount type.
	errorBarsYFormat2->SetErrorBarvType(ErrorValueType::StandardError);

	//Set value.
	errorBarsYFormat2->SetErrorBarVal(0.3);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
