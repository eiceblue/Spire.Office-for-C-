#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_HtmlFile1.html";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"HtmlToImage.png";


	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str(), FileFormat::Html, XHTMLValidationType::None);

	//Save to image in the default format of png.
#if defined(SKIASHARP)
	SkiaSharp::SKintrusive_ptr<Image> images = document->SaveToImages(0, ImageType::Bitmap);
	Fileintrusive_ptr<Stream> fileStream = new FileStream(outputFile.c_str(), FileMode::Create, FileAccess::Write);
	images->Encode(SkiaSharp::SKEncodedFREE_IMAGE_FORMAT::FIF_PNG, 100)->SaveTo(fileStream);
	fileStream->Flush();
#else

	intrusive_ptr<Stream> imageStream = document->SaveImageToStreams(0, ImageType::Bitmap);
	//Obtain image data in the default format of png,you can use it to convert other image format
	std::vector<byte> imageData = imageStream->ToArray();
	std::ofstream outFile(outputFile, std::ios::binary);
	if (outFile.is_open())
	{
		outFile.write(reinterpret_cast<const char*>(imageData.data()), imageData.size());
		outFile.close();
	}
#endif
	document->Close();
	imageStream->Dispose();
}
