
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_3.pptx";
	wstring outputFile = OUTPUTPATH"RemoveChartFromPptSlide.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide from the document.
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Remove chart from the slide.
	for (int i = slide->GetShapes()->GetCount() - 1; i >= 0; i--)
	{
		intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(i);
		intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(shape);
		if (chart != nullptr)
		{
			slide->GetShapes()->Remove(shape);
		}
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
