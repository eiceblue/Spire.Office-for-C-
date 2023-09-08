#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToImage.png";

	//Create word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Save doc file.
#if defined(SKIASHARP)
	SkiaSharp::SKintrusive_ptr<Image> images = document->SaveToImages(0, ImageType::Metafile);
	Fileintrusive_ptr<Stream> fileStream = new FileStream(outputFile.c_str(), FileMode::Create, FileAccess::Write);
	images->Encode(SkiaSharp::SKEncodedFREE_IMAGE_FORMAT::FIF_PNG, 100)->SaveTo(fileStream);
	fileStream->Flush();
#else

	intrusive_ptr<Stream> imageStream = document->SaveImageToStreams(0, ImageType::Bitmap);
	//Obtain image data in the default format of png,you can use it to convert other image format
	std::vector<byte> imgBytes = imageStream->ToArray();
	std::ofstream outFile(outputFile, std::ios::binary);
	if (outFile.is_open())
	{
		outFile.write(reinterpret_cast<const char*>(imgBytes.data()), imgBytes.size());
		outFile.close();
	}
#endif
	document->Close();
	imageStream->Dispose();

	document->Close();
	imageStream->Dispose();
}
