#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"PPTHasHeader.pptx";
	wstring outputFile = OUTPUTPATH"ManageNoteMasterHeaderFooter.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Set the note Masters header and footer
	intrusive_ptr<INoteMasterSlide> noteMasterSlide = ppt->GetNotesMaster();
	if (noteMasterSlide != nullptr)
	{
		intrusive_ptr<ShapeCollection> shapes = noteMasterSlide->GetShapes();
		for (int i = 0; i < shapes->GetCount(); i++)
		{
			intrusive_ptr<IShape> shape = shapes->GetItem(i);
			if (shape->GetPlaceholder() != nullptr)
			{
				if (shape->GetPlaceholder()->GetType() == PlaceholderType::Header)
				{
					(Object::Dynamic_cast<IAutoShape>(shape))->GetTextFrame()->SetText(L"change the header by Spire");
				}
				if (shape->GetPlaceholder()->GetType() == PlaceholderType::Footer)
				{
					(Object::Dynamic_cast<IAutoShape>(shape))->GetTextFrame()->SetText(L"change the footer by Spire");
				}
			}
		}
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
