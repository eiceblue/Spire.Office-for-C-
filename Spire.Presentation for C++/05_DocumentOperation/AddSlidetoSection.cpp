#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Section.pptx";
	wstring outputFile = OUTPUTPATH"AddSlidetoSection.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Add a new shape to the PPT document
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(200, 50, 300, 100));

	//Create a new section and copy the first slide to it
	intrusive_ptr<Section> NewSection = ppt->GetSectionList()->Append(L"New Section");
	NewSection->Insert(0, ppt->GetSlides()->GetItem(0));

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
