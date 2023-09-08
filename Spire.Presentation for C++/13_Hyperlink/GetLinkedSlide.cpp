#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"linkedSlide.pptx";
	wstring outputFile = OUTPUTPATH"GetLinkedSlide.txt";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the second slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(1);

	//Get the first shape of the second slide
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(slide->GetShapes()->GetItem(0));
	wofstream outFile(outputFile);
	//Get the linked slide index
	if (shape->GetClick()->GetActionType() == HyperlinkActionType::GotoSlide)
	{
		intrusive_ptr<ISlide> targetSlide = shape->GetClick()->GetTargetSlide();
		outFile << "Linked slide number = " << targetSlide->GetSlideNumber() << endl;
	}
	outFile.close();
}
