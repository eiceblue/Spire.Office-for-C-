#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"PPTSample_1.pptx";
	wstring outputFile = OUTPUTPATH"RemoveUnusedLayoutMaster.pptx";

	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Create an array list
	std::vector<intrusive_ptr<IActiveSlide>> list;
	for (int i = 0; i < ppt->GetSlides()->GetCount(); i++)
	{
		//Get the layout used by slide
		intrusive_ptr<IActiveSlide> layout= Object::Dynamic_cast<IActiveSlide>(ppt->GetSlides()->GetItem(i)->GetLayout().get());
		list.push_back(layout);
	}

	//Loop through masters and layouts
	for (int i = 0; i < ppt->GetMasters()->GetCount(); i++)
	{
		intrusive_ptr<IMasterLayouts> masterlayouts = ppt->GetMasters()->GetItem(i)->GetLayouts();
		for (int j = masterlayouts->GetCount() - 1; j >= 0; j--)
		{
			intrusive_ptr<IActiveSlide> master = Object::Dynamic_cast<IActiveSlide>(masterlayouts->GetItem(j));
			if (!(std::find(list.begin(), list.end(), master) != list.end()))
			{
				//Remove unused layout
				masterlayouts->RemoveMasterLayout(j);
			}
		}

	}
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	
}
