#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"HasPromptText.pptx";
	wstring outputFile = OUTPUTPATH"EditPromptText.pptx";

	//Load a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	// Iterate through the slide
	for (int s = 0; s < presentation->GetSlides()->GetItem(0)->GetShapes()->GetCount(); s++) {
		intrusive_ptr<IShape> shape = presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(s);
		if (shape->GetPlaceholder() != nullptr && Object::Dynamic_cast<IAutoShape>(shape) != nullptr)
		{
			wstring text = L"";
			// Set the text of the title
			if (shape->GetPlaceholder()->GetType() == PlaceholderType::CenteredTitle)
			{
				text = L"custom title create by Spire";
			}
			// Set text of the subtitle.
			else if (shape->GetPlaceholder()->GetType() == PlaceholderType::Subtitle)
			{
				text = L"custom subtitle create by Spire";
			}

			(Object::Dynamic_cast<IAutoShape>(shape))->GetTextFrame()->SetText(text.c_str());
		}
	}
	//Save the file
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}

