
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChartSample1.pptx";
	wstring outputFile = OUTPUTPATH"AddCustomErrorBars.pptx";

	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	//Get the column chart on the first slide and set chart title.
	intrusive_ptr<IChart> bubbleChart = Object::Dynamic_cast<IChart>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Get X error bars of the first chart series
	intrusive_ptr<IErrorBarsFormat> errorBarsXFormat = bubbleChart->GetSeries()->GetItem(0)->GetErrorBarsXFormat();
	//Specify error amount type as custom error bars
	errorBarsXFormat->SetErrorBarvType(ErrorValueType::CustomErrorBars);
	//Set the minus and plus value of the X error bars
	errorBarsXFormat->SetMinusVal(0.5);
	errorBarsXFormat->SetPlusVal(0.5);

	//Get Y error bars of the first chart series
	intrusive_ptr<IErrorBarsFormat> errorBarsYFormat = bubbleChart->GetSeries()->GetItem(0)->GetErrorBarsYFormat();
	//Specify error amount type as custom error bars
	errorBarsYFormat->SetErrorBarvType(ErrorValueType::CustomErrorBars);
	//Set the minus and plus value of the Y error bars
	errorBarsYFormat->SetMinusVal(1.0);
	errorBarsYFormat->SetPlusVal(1.0);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
