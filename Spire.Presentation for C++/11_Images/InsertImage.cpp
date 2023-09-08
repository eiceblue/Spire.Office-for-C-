#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"InsertImage.pptx";
	wstring outputFile = OUTPUTPATH"InsertImage.pptx";
	

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//EMF file path
	wstring ImageFile = DATAPATH"InsertImage.png";

	//Define image size
	intrusive_ptr<RectangleF> rect = new RectangleF(ppt->GetSlideSize()->GetSize()->GetWidth() / 2 - 280, 140, 120, 120);

	//Append the EMF in slide
	intrusive_ptr<IEmbedImage> image = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	image->GetLine()->SetFillType(FillFormatType::None);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}

