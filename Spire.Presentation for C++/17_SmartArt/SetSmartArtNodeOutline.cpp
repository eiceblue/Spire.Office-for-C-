#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SmartArtLinklineOutline.pptx";
	wstring outputFile = OUTPUTPATH"SetSmartArtNodeOutline.pptx";

	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());
	intrusive_ptr<ISmartArt> smartArt = Object::Dynamic_cast<ISmartArt>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));
	int count = smartArt->GetNodes()->GetCount();
	intrusive_ptr<ISmartArtNode> node;
	//Loop through all nodes
	for (int i = 0; i < count; i++)
	{
		node = smartArt->GetNodes()->GetItem(i);
		//Set the fill format type
		node->GetLine()->SetFillType(FillFormatType::Solid);
		//Set the line style
		node->GetLine()->SetStyle(TextLineStyle::ThinThin);
		//Set the line color
		node->GetLine()->GetSolidFillColor()->SetColor(Color::GetRed());
		//Set the line width
		node->GetLine()->SetWidth(2);
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
				
}
