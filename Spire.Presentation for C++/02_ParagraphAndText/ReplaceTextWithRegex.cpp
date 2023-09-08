#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SomePresentation.pptx";
	wstring outputFile = OUTPUTPATH"ReplaceTextWithRegex.pptx";

	//Create Presentation
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load file
	presentation->LoadFromFile(inputFile.c_str());

	//Regex for all words
	intrusive_ptr<Regex> regex = new Regex(DATAPATH"(\\d+.\\d+|\\w+)");

	//New string value
	std::string newvalue = "This is the test!";

	//Loop and replace
	for (int s = 0; s < presentation->GetSlides()->GetCount(); s++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(s);
		for (int i = 0; i < slide->GetShapes()->GetCount(); i++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(i);
			//shape->ReplaceTextWithRegex(regex, newvalue);     //Regex ²»Ö§³Ö
		}
	}

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

