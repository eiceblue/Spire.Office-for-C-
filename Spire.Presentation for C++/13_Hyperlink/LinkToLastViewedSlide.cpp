#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"LastViewedSlide.pptx";
	wstring outputFile = OUTPUTPATH"LinkToLastViewedSlide.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	ppt->LoadFromFile(inputFile.c_str());

	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);
	//Draw a shape
	intrusive_ptr<IAutoShape> autoShape = slide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(100, 100, 100, 100));
	//Link to last viewed slide show
	autoShape->SetClick(ClickHyperlink::GetLastVievedSlide());

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
