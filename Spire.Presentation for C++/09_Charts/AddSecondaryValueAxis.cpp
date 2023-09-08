
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_2.pptx";
	wstring outputFile = OUTPUTPATH"AddSecondaryValueAxisToChart.pptx";


	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the column chart on the first slide and set chart title.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Add a secondary axis to display the value of Series 3.
	chart->GetSeries()->GetItem(2)->SetUseSecondAxis(true);

	//Set the grid line of secondary axis as invisible.
	chart->GetSecondaryValueAxis()->GetMajorGridTextLines()->SetFillType(FillFormatType::None);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
