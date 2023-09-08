#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"TwoColumns.pptx";
	wstring outputFile = OUTPUTPATH"SetRightToLeftColumns.pptx";

	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	//Get the second shape
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));
	//Set columns style to right-to-left
	shape->GetTextFrame()->SetRightToLeftColumns(true);
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}

