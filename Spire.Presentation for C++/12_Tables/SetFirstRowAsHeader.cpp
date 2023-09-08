#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"NormalTable.pptx";
	wstring outputFile = OUTPUTPATH"SetFirstRowAsHeader.pptx";

	intrusive_ptr<ITable> table = nullptr;

	//Load a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (Object::Dynamic_cast<ITable>(shape) != nullptr)
		{
			table = Object::Dynamic_cast<ITable>(shape);
		}

	}
	table->SetFirstRow(true);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
		
}
