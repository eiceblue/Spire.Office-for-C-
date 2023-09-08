#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_1.pptx";
	wstring outputFile = DATAPATH"SetBordersForExistingTable.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the table from the first slide of the sample document.
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);
	intrusive_ptr<ITable> table = Object::Dynamic_cast<ITable>(slide->GetShapes()->GetItem(0));

	//Set the border type as Inside and the border color as blue.
	table->SetTableBorder(TableBorderType::Inside, 1, Color::GetBlue());

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
			
}
