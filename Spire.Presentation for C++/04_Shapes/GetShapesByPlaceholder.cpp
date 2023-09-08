#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"GetShapesByPlaceholder.pptx";
	wstring outputFile = OUTPUTPATH"GetShapesByPlaceholder.pptx";

	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	intrusive_ptr<Placeholder> placeholder = ppt->GetSlides()->GetItem(1)->GetShapes()->GetItem(0)->GetPlaceholder();
	//Get Shapes by Placeholder
	std::vector<intrusive_ptr<IShape>> shapes = ppt->GetSlides()->GetItem(1)->GetPlaceholderShapes(placeholder);
	wstring text = L"";
	//Iterate over all the shapes
	for (int i = 0; i < shapes.size(); i++)
	{
		//If shape is IAutoShape
		if (Object::Dynamic_cast<IAutoShape>(shapes[i]) != nullptr)
		{
			intrusive_ptr<IAutoShape> autoShape = Object::Dynamic_cast<IAutoShape>(shapes[i]);
			if (autoShape->GetTextFrame() != nullptr)
			{
				//text += autoShape->GetTextFrame()->GetText() + "\\r\\n";
				text += autoShape->GetTextFrame()->GetText();
			}
		}
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

