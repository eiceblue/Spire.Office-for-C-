#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"BlankSample_N.pptx";
	wstring outputFile = OUTPUTPATH"GetSlideByIndexOrID.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get slide by index 0
	intrusive_ptr<ISlide> slide1 = presentation->GetSlides()->GetItem(0);
	//Append a shape in the slide
	intrusive_ptr<IAutoShape> shape1 = slide1->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(100, 100, 200, 100));
	//Add text in the shape
	shape1->GetTextFrame()->SetText(L"Get slide by index");

	//Get slide by slide ID
	intrusive_ptr<ISlide> slide2 = presentation->FindSlide(static_cast<int>(presentation->GetSlides()->GetItem(1)->GetSlideID()));
	//Append a shape in the slide
	intrusive_ptr<IAutoShape> shape2 = slide2->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(100, 100, 200, 100));
	//Add text in the shape
	shape2->GetTextFrame()->SetText(L"Get slide by slide id");

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	
}
