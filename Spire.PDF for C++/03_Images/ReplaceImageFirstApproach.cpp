#include "../pch.h"

using namespace std;
using namespace Spire::Pdf;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path+L"ReplaceImage.pdf";
	wstring inputImg = input_path+L"E-iceblueLogo.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"ReplaceImageFirstApproach.pdf";

	//Create a pdf document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile(inputFile.c_str());
	//Get the first page.
	intrusive_ptr<PdfPageBase> page = pdf->GetPages()->GetItem(0);
	//Get images of the first page.
	std::vector<intrusive_ptr<PdfImageInfo>> imageInfo = page->GetImagesInfo();
	intrusive_ptr<PdfImage> image = new PdfImage();
	image = image->FromFile(inputImg.c_str());
	page->ReplaceImage(0, image);
	//Save the document
	pdf->SaveToFile(outputFile.c_str());

	pdf->Close();
}









