#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"ImageWatermarkFirst.pdf";
	wstring inputFile = DATAPATH"ImageWaterMark.pdf";

	//Load the Pdf from disk
	intrusive_ptr<PdfDocument> document = new PdfDocument();
	document->LoadFromFile(inputFile.c_str());
	//Get the first page
	intrusive_ptr<PdfPageBase> page = document->GetPages()->GetItem(0);
	//Load image          
	wstring inputFile_i = DATAPATH"Background.png";
	//intrusive_ptr<Image> img = Image::FromFile(inputFile1.c_str());
	intrusive_ptr<Stream> img = new Stream(inputFile_i.c_str());
	//Set background image

	page->SetBackgroundImage(img);

	//Save the document
	document->SaveToFile(outputFile.c_str());
	document->Close();
}