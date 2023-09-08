
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;


const wstring chartTypeToString(ChartBaseUnitType chartType)
{
	switch (chartType)
	{
	case Spire::Presentation::ChartBaseUnitType::Days:
		return L"Days";
		break;
	case Spire::Presentation::ChartBaseUnitType::Months:
		return L"Months";
		break;
	case Spire::Presentation::ChartBaseUnitType::Years:
		return L"Years";
		break;
	case Spire::Presentation::ChartBaseUnitType::Auto:
		return L"Auto";
		break;
	default:
		break;
	}
}

int main()
{
	wstring inputFile = DATAPATH"ChartSample2.pptx";
	wstring outputFile = OUTPUTPATH"GetValuesAndUnitFromAxis.txt";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	wofstream desFile(outputFile, ios::out);
	//Get unit from primary category axis
	float MajorUnit = chart->GetPrimaryCategoryAxis()->GetMajorUnit();
	ChartBaseUnitType type = chart->GetPrimaryCategoryAxis()->GetMajorUnitScale();

	desFile << MajorUnit << endl;
	desFile << chartTypeToString(type).c_str() << endl;

	//Get values from primary value axis
	float minValue = chart->GetPrimaryValueAxis()->GetMinValue();
	float maxValue = chart->GetPrimaryValueAxis()->GetMaxValue();

	desFile << minValue << endl;
	desFile << maxValue << endl;

	desFile.close();
}
