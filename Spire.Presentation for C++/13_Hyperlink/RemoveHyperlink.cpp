#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_5.pptx";
	wstring outputFile = OUTPUTPATH"RemoveHyperlink.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	ppt->LoadFromFile(inputFile.c_str());

	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);
	//Find the hyperlinks you want to edit.
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(slide->GetShapes()->GetItem(0));

	//Set the ClickAction property into null to remove the hyperlink.
	shape->GetTextFrame()->GetTextRange()->SetClickAction(ClickHyperlink::GetNoAction());

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
