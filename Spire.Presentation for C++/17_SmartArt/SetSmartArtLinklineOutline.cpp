#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SmartArtLinklineOutline.pptx";
	wstring outputFile = OUTPUTPATH"SetSmartArtLinklineOutline.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load document from disk
	ppt->LoadFromFile(inputFile.c_str());
	intrusive_ptr<ISmartArt> smartArt = Object::Dynamic_cast<ISmartArt>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	int count = smartArt->GetNodes()->GetCount();
	intrusive_ptr<ISmartArtNode> node;
	//Loop through all smartArts
	for (int i = 0; i < count; i++)
	{
		node = smartArt->GetNodes()->GetItem(i);
		//Set the line type
		node->GetLinkLine()->SetFillType(FillFormatType::Solid);
		//Set the line color
		node->GetLinkLine()->GetSolidFillColor()->SetColor(Color::GetRed());
		//Set the line width
		node->GetLinkLine()->SetWidth(2);
		//Set the line DashStyle
		node->GetLinkLine()->SetDashStyle(LineDashStyleType::SystemDash);
	}
	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
		
}
