#include "pch.h"

using namespace Spire::Common;
using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"EmbedImagesInHTML.pdf";
	wstring outputFile = OUTPUTPATH"EmbedImages.html";

	//Create a pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	doc->LoadFromFile(inputFile.c_str());

	//Set the convertion option to embed image in html
	doc->GetConvertOptions()->SetPdfToHtmlOptions(true, true);

	//Save the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::HTML);
	doc->Close();
}
