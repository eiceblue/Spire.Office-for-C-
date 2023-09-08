
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SetAxisType.pptx";
	wstring outputFile = OUTPUTPATH"SetAxisType.pptx";

	// Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));

	chart->GetPrimaryCategoryAxis()->SetAxisType(AxisType::DateAxis);
	chart->GetPrimaryCategoryAxis()->SetMajorUnitScale(ChartBaseUnitType::Months);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
