#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SetTableStyle.pptx";
	wstring outputFile = OUTPUTPATH"SetTableStyle.pptx";

	//Creat a ppt document and load file
	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get tbe table
	intrusive_ptr<ITable> table = nullptr;
	for (int s = 0; s < ppt->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		table = Object::Dynamic_cast<ITable>(shape);
		if (table != nullptr)
		{
			//Set the table style from TableStylePreset and apply it to selected table
			table->SetStylePreset(TableStylePreset::MediumStyle1Accent2);
		}
	}
	//Save the file
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
		
}
