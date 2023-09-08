#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetImageFrameFormat.pptx";
	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Load an image
	wstring imageFile = DATAPATH"iceblueLogo.png";
	intrusive_ptr<Stream> stream = new Stream(imageFile.c_str());
	intrusive_ptr<IImageData> imageData = ppt->GetImages()->Append(stream);

	//Add the image in document
	intrusive_ptr<RectangleF> rect = new RectangleF(100, 100, imageData->GetWidth() / 2, imageData->GetHeight() / 2);
	intrusive_ptr<IEmbedImage> pptImage = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, imageData, rect);

	//Set the formatting of the image frame
	pptImage->GetLine()->GetFillFormat()->SetFillType(FillFormatType::Solid);
	pptImage->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetLightBlue());
	pptImage->GetLine()->SetWidth(5);
	pptImage->SetRotation(-45);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

