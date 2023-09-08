#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddImageInMaster.pptx";
	wstring outputFile = OUTPUTPATH"AddImageInMaster.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the master collection
	intrusive_ptr<IMasterSlide> master = presentation->GetMasters()->GetItem(0);

	//Append image to slide master
	wstring image = DATAPATH"Logo.png";
	intrusive_ptr<RectangleF> rff = new RectangleF(40, 40, 90, 90);
	intrusive_ptr<IEmbedImage> pic = master->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, image.c_str(), rff);
	pic->GetLine()->GetFillFormat()->SetFillType(FillFormatType::None);

	//Add new slide to presentation
	presentation->GetSlides()->Append();

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
			
}
