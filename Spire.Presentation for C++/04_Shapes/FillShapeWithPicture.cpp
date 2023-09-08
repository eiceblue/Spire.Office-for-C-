#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"FillShapeWithPicture.pptx";
	wstring outputFile = OUTPUTPATH"FillShapeWithPicture.pptx";

	//Load a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first shape and set the style to be Gradient
	intrusive_ptr<IAutoShape> shape = Object::Dynamic_cast<IAutoShape>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Fill the shape with picture
	wstring picUrl = DATAPATH"backgroundImg.png";
	shape->GetFill()->SetFillType(FillFormatType::Picture);
	shape->GetFill()->GetPictureFill()->GetPicture()->SetUrl(picUrl.c_str());
	shape->GetFill()->GetPictureFill()->SetFillType(PictureFillType::Stretch);

	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}

