
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChartSample2.pptx";
	wstring outputFile = OUTPUTPATH"ChangeSeriesName.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get chart on the first slide
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Get the ranges of series label 
	intrusive_ptr<CellRanges> cr = chart->GetSeries()->GetSeriesLabel();

	//Change the value
	cr->GetItem(0)->SetText(L"Changed series name");

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
