#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"EmbedZipIntoPPT.pptx";
	wstring inputFile_z = DATAPATH"test.zip";
	wstring inputFile_I = DATAPATH"icon.png";
	wstring outputFile = OUTPUTPATH"EmbedZipIntoPPT.pptx";
	

	//Create a Presentaion document
	intrusive_ptr<Presentation> ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Load a zip object
	intrusive_ptr<Stream> zipStream = new Stream(inputFile_z.c_str());

	intrusive_ptr<RectangleF> rec = new RectangleF(80, 60, 100, 100);

	//Insert the zip object to presentation
	intrusive_ptr<Spire::Presentation::IOleObject> ole = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendOleObject(L"zip", zipStream, rec);
	ole->SetProgId(L"Package");
	intrusive_ptr<Stream> image = new Stream(inputFile_I.c_str());
	intrusive_ptr<IImageData> oleImage = ppt->GetImages()->Append(image);
	ole->GetSubstituteImagePictureFillFormat()->GetPicture()->SetEmbedImage(oleImage);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}