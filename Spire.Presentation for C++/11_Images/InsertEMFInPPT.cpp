#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"BlankSample_N.pptx";
	wstring outputFile = OUTPUTPATH"InsertEMFInPPT.pptx";
	

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//EMF file path
	wstring ImageFile = DATAPATH"InsertEMF.emf";

	//Define image size
	intrusive_ptr<RectangleF> rect = new RectangleF(100, 100, 719 / 1.5, 539 / 1.5);
	//Append the EMF in slide
	intrusive_ptr<IEmbedImage> image = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	image->GetLine()->SetFillType(FillFormatType::None);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

