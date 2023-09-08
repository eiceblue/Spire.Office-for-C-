#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"GroupShapes.pptx";
	wstring outputFile = OUTPUTPATH"UngroupShapes.pptx";

	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	intrusive_ptr<GroupShape> groupShape = Object::Dynamic_cast<GroupShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	//Ungroup the shapes
	ppt->GetSlides()->GetItem(0)->Ungroup(groupShape);
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

