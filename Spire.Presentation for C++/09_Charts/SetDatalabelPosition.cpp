#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChartSample2.pptx";
	wstring outputFile = OUTPUTPATH"SetDatalabelPosition.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Add data label
	intrusive_ptr<ChartDataLabel> label = chart->GetSeries()->GetItem(0)->GetDataLabels()->Add();
	//Set the position of the label
	label->SetX(0.1f);
	label->SetY(0.1f);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
