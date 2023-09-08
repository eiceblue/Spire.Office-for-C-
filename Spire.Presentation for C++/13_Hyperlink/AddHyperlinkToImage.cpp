#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_5.pptx";
	wstring outputFile = OUTPUTPATH"AddHyperlinkToImage.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide.
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Add image to slide.
	intrusive_ptr<RectangleF> rect = new RectangleF(480, 350, 160, 160);
	std::wstring inputFile1 = DATAPATH"Logo1.png";
	intrusive_ptr<IEmbedImage> image = slide->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, inputFile1.c_str(), rect);

	//Add hyperlink to the image.
	intrusive_ptr<ClickHyperlink> hyperlink = new ClickHyperlink(L"https://www.e-iceblue.com");
	image->SetClick(hyperlink);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
