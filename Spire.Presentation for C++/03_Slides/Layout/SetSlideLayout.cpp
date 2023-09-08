#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetSlideLayout.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Remove the first slide
	ppt->GetSlides()->RemoveAt(0);

	//Append a slide and set the layout for slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->Append(SlideLayoutType::Title);

	//Add content for Title and Text
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(slide->GetShapes()->GetItem(0));
	shape->GetTextFrame()->SetText(L"Hello Wolrd! -> This is title");

	shape = Object::Dynamic_cast<IAutoShape>(slide->GetShapes()->GetItem(1));
	shape->GetTextFrame()->SetText(L"E-iceblue Support Team -> This is content");

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
					
}
