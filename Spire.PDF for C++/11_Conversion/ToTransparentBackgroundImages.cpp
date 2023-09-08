#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"ToTransparentBackgroundImages.png";
	wstring inputFile = DATAPATH"ToTransparentBackgroundImages.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	doc->GetConvertOptions()->SetPdfToImageOptions(0);
	intrusive_ptr<Stream> image = doc->SaveAsImage(0, PdfImageType::Bitmap);
	image->Save(outputFile.c_str());
}
