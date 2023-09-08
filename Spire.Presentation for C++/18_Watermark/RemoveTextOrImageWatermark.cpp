#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"RemoveTextAndImageWatermarks.pptx";
	wstring outputFile = OUTPUTPATH"RemoveTextAndImageWatermarks.pptx";

	//Create a PowerPoint document.
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Remove text watermark by removing the shape which contains the text string "E-iceblue".
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		for (int j = 0; j < presentation->GetSlides()->GetItem(i)->GetShapes()->GetCount(); j++)
		{
			intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape > (presentation->GetSlides()->GetItem(i)->GetShapes()->GetItem(j));

			if (shape != nullptr)
			{
				std::wstring temp = shape->GetTextFrame()->GetText();
				std::string::size_type pos = temp.find(L"E-iceblue");
				if (pos != string::npos)
				{
					presentation->GetSlides()->GetItem(i)->GetShapes()->Remove(shape);
				}
			}
		}
	}
	//Remove image watermark.
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		presentation->GetSlides()->GetItem(i)->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::None);
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	
}


