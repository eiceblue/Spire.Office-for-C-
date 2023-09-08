#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_7.pptx";
	wstring outputFile = OUTPUTPATH"ResetPositionOfDateTimeAndSlideNumber.pptx";

	//Create a PowerPoint document.
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get the first slide from the sample document.
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);

	for (int s = 0; s < slide->GetShapes()->GetCount(); s++)
	{
		intrusive_ptr<IShape> shapeToMove = slide->GetShapes()->GetItem(s);
		//Reset the position of the slide number to the left.
		wstring temp = shapeToMove->GetName();
		wstring::size_type pos = temp.find(L"Slide Number Placeholder");
		wstring temp1 = shapeToMove->GetName();
		wstring::size_type pos1 = temp.find(L"Date Placeholder");
		if (pos != string::npos)
		{
			shapeToMove->SetLeft(0);
		}
		else if (pos1 != string::npos)
		{
			//Reset the position of the date time to the center.
			shapeToMove->SetLeft(presentation->GetSlideSize()->GetSize()->GetWidth() / 2);

			//Reset the date time display style.
			(Object::Dynamic_cast<IAutoShape>(shapeToMove))->GetTextFrame()->GetTextRange()->GetParagraph()->SetText((DateTime::GetNow())->ToString());
			(Object::Dynamic_cast<IAutoShape>(shapeToMove))->GetTextFrame()->SetIsCentered(true);
		}
	}
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	
}

