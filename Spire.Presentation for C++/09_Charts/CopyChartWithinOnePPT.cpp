
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_2.pptx";
	wstring outputFile = OUTPUTPATH"CopyChartWithinOnePPT.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get chart on the first slide
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Copy the chart from the first slide to the specified location of the second slide within the same document.
	intrusive_ptr<ISlide> slide1 = ppt->GetSlides()->Append();
	slide1->GetShapes()->CreateChart(chart, new RectangleF(100, 100, 500, 300), 0);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
