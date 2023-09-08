#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"WordToEmf.emf";

	//Create a Word document.
	intrusive_ptr<Document> document = new Document();
	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str(), FileFormat::Docx);
	//Convert the first page of document to image.
#if defined(SKIASHARP)
	SkiaSharp::SKintrusive_ptr<Image> images = document->SaveToImages(0, ImageType::Metafile);
	Fileintrusive_ptr<Stream> fileStream = new FileStream(outputFile.c_str(), FileMode::Create, FileAccess::Write);
	images->Encode(SkiaSharp::SKEncodedFREE_IMAGE_FORMAT::FIF_PNG, 100)->SaveTo(fileStream);
	fileStream->Flush();
#else
				//intrusive_ptr<Image> image = document->SaveToImages(0, ImageType::Metafile);
				//Save the file.
				//image->Save(outputFile.c_str(), ImageFormat::GetEmf());
	intrusive_ptr<Stream> imageStream = document->SaveImageToStreams(0, ImageType::Bitmap);
	//Obtain image data in the default format of png,you can use it to convert other image format
	std::vector<byte> imageData = imageStream->ToArray();
	//TestUtil::WriteBytesToImage(imageData, outputFile,FREE_IMAGE_FORMAT::FIF_PNG);
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