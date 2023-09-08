#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Titles.pptx";
	wstring outputFile = OUTPUTPATH"GetAllTitles.txt";

	wofstream outFile(outputFile, ios::out);

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Instantiate a list of IShape objects
	std::vector<intrusive_ptr<IShape>> shapelist;
	//Loop through all sildes and all shapes on each slide
	for (int s = 0; s < ppt->GetSlides()->GetCount(); s++)
	{
		intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(s);
		for (int i = 0; i < slide->GetShapes()->GetCount(); i++) {
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(i);
			if (shape->GetPlaceholder() != nullptr)
			{
				//Get all titles
				switch (shape->GetPlaceholder()->GetType())
				{
				case PlaceholderType::Title:
					shapelist.push_back(shape);
					break;
				case PlaceholderType::CenteredTitle:
					shapelist.push_back(shape);
					break;
				case PlaceholderType::Subtitle:
					shapelist.push_back(shape);
					break;
				}
			}
		}
	}

	//Loop through the list and get the inner text of all shapes in the list
	outFile << "Below are all the obtained titles:" << endl;
	for (int i = 0; i < shapelist.size(); i++)
	{
		intrusive_ptr<IAutoShape> shape1 = Object::Dynamic_cast<IAutoShape>(shapelist[i]);
		outFile << shape1->GetTextFrame()->GetText() << endl;
	}

	//Save to the Text file
	outFile.close();
}

