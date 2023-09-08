#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddImageInTableCell.pptx";
	wstring outputFile = OUTPUTPATH"AddImageInTableCell.pptx";

	//Load a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first shape
	intrusive_ptr<ITable> table = Object::Dynamic_cast<ITable>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Load the image and insert it into table cell

	intrusive_ptr<Stream> stream = new Stream(DATAPATH"PresentationIcon.png");
	intrusive_ptr<IImageData> pptImg = ppt->GetImages()->Append(stream);
	stream->Close();

	table->GetItem(1, 1)->GetFillFormat()->SetFillType(FillFormatType::Picture);
	table->GetItem(1, 1)->GetFillFormat()->GetPictureFill()->GetPicture()->SetEmbedImage(pptImg);
	table->GetItem(1, 1)->GetFillFormat()->GetPictureFill()->SetFillType(PictureFillType::Stretch);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
