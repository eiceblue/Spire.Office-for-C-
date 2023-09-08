#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"IsTextboxSample.pptx";
	wstring outputFile = OUTPUTPATH"IsTextBox.txt";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);

	for (int l = 0; l < ppt->GetSlides()->GetCount(); l++)
	{
		intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(l);
		for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(s);
			if (Object::Dynamic_cast<IAutoShape>(shape) != nullptr)
			{
				//Judge if the shape is textbox
				bool isTextbox = shape->GetIsTextBox();
				outFile << (isTextbox ? "shape is text box" : "shape is not text box") << endl;
			}
		}
	}
	outFile.close();
}

