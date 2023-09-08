#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetImageTransparency.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Load an image
	wstring imageFile = DATAPATH"Logo.png";
	ifstream inputf(imageFile.c_str(), std::ios::in | std::ios::binary);
	intrusive_ptr<Stream> stream = new Stream(inputf);
	intrusive_ptr<IImageData> imageData = ppt->GetImages()->Append(stream);

	//Add the image in document
	intrusive_ptr<RectangleF> rect = new RectangleF(200, 100, imageData->GetWidth(), imageData->GetHeight());
	//Add a shape
	intrusive_ptr<IAutoShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, rect);
	shape->GetLine()->SetFillType(FillFormatType::None);
	//Fill shape with image
	shape->GetFill()->SetFillType(FillFormatType::Picture);
	shape->GetFill()->GetPictureFill()->GetPicture()->SetUrl(imageFile.c_str());
	shape->GetFill()->GetPictureFill()->SetFillType(PictureFillType::Stretch);
	//Set transparency on image
	shape->GetFill()->GetPictureFill()->GetPicture()->SetTransparency(50);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

