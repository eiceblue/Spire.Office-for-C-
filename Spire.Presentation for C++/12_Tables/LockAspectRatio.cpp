#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Table.pptx";
	wstring outputFile = OUTPUTPATH"LockAspectRatio.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load PPT file from disk
	presentation->LoadFromFile(inputFile.c_str());
	//Get the first slide
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		//Verify if it is table
		intrusive_ptr<ITable> table = Object::Dynamic_cast<ITable >(shape);
		if (table != nullptr)
		{
			//Lock aspect ratio
			table->GetShapeLocking()->SetAspectRatioProtection(true);
		}
	}

	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
			
}
