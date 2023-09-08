
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile_1 = DATAPATH"Template_Ppt_2.pptx";
	wstring inputFile_2 = DATAPATH"Template_Ppt_1.pptx";
	wstring outputFile = OUTPUTPATH"CopyChartBetweenPptFiles.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile_1.c_str());

	//Get chart on the first slide
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Create a PPT document
	intrusive_ptr<Presentation> presentation2 = new Presentation();
	//Load the file from disk.
	presentation2->LoadFromFile(inputFile_2.c_str());

	//Copy chart from the first document to the second document.
	presentation2->GetSlides()->Append();
	presentation2->GetSlides()->GetItem(1)->GetShapes()->CreateChart(chart, new RectangleF(100, 100, 500, 300), -1);

	//Save to file.
	presentation2->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
