#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"EmbedExcelAsOLE.xlsx";
	wstring outputFile = OUTPUTPATH"EmbedExcelAsOLE.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	wstring imageFile = DATAPATH"EmbedExcelAsOLE.png";
	intrusive_ptr<IImageData> oleImage = ppt->GetImages()->Append(new Stream(imageFile.c_str()));

	intrusive_ptr<RectangleF> rec = new RectangleF(80, 60, oleImage->GetWidth(), oleImage->GetHeight());
	//Insert an OLE object to presentation based on the Excel data
	intrusive_ptr<Spire::Presentation::IOleObject> oleObject = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendOleObject(L"excel", new Stream(inputFile.c_str()), rec);
	oleObject->GetSubstituteImagePictureFillFormat()->GetPicture()->SetEmbedImage(oleImage);
	oleObject->SetProgId(L"Excel.Sheet.12");

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}

