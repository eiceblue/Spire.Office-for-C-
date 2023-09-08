#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ExtractText.pptx";
	wstring outputFile = OUTPUTPATH"ExtractText.txt";

	//Create a PPT document and load file
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	wofstream desFile(outputFile, ios::out);
	//Foreach the slide and extract text
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(i);
		for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(s);
			if (Object::Dynamic_cast<IAutoShape>(shape) != nullptr)
			{
				for (int t = 0; t < (Object::Dynamic_cast<IAutoShape>(shape))->GetTextFrame()->GetParagraphs()->GetCount(); t++)
				{
					intrusive_ptr<TextParagraph> tp =
						(Object::Dynamic_cast<IAutoShape>(shape))->GetTextFrame()->GetParagraphs()->GetItem(t);
					desFile << tp->GetText() << endl;
				}
			}
		}
	}
	desFile.close();


}

