#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_1.pptx";
	wstring outputFile = DATAPATH"SplitSpecificTableCell.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide.
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	//Get the table.
	intrusive_ptr<ITable> table = Object::Dynamic_cast<ITable>(slide->GetShapes()->GetItem(0));

	//Split cell [1, 2) into 3 rows and 2 columns.
	table->GetItem(1, 2)->Split(3, 2);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
		
}
