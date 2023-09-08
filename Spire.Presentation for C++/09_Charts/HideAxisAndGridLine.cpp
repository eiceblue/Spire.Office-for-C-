
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChartSample2.pptx";
	wstring outputFile = OUTPUTPATH"HideAxisAndGridLine.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Hide axis
	chart->GetPrimaryCategoryAxis()->SetIsVisible(false);
	chart->GetPrimaryValueAxis()->SetIsVisible(false);

	//Remove gridline
	chart->GetPrimaryValueAxis()->GetMajorGridTextLines()->SetFillType(FillFormatType::None);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
