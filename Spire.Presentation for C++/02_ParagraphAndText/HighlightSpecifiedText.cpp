#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	
	wstring inputFile = DATAPATH"SomePresentation.pptx";
	wstring outputFile = OUTPUTPATH"HighlightSpecifiedText.pptx";

	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get the specified shape
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));

	intrusive_ptr<TextHighLightingOptions> options = new TextHighLightingOptions();
	options->SetWholeWordsOnly(true);
	options->SetCaseSensitive(true);

	shape->GetTextFrame()->HighLightText(L"Spire", Color::GetYellow(), options);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

