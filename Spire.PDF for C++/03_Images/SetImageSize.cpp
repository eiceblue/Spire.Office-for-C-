#include "../pch.h"

using namespace std;
using namespace Spire::Pdf;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ChartImage.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetImageSize.pdf";

	//Create a pdf document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	//Create one page
	intrusive_ptr<PdfPageBase> page = new PdfPageBase();
	page = pdf->GetPages()->Add();
	intrusive_ptr<PdfImage> image = new PdfImage();
	image = image->FromFile(inputFile.c_str());
	//Set the width and height of image
	float width = image->GetWidth() * 0.75f;
	float height = image->GetHeight() * 0.75f;
	//Define a position to draw image
	float x = (page->GetCanvas()->GetClientSize()->GetWidth() - width) / 2;
	float y = 60.0f;
	//Draw image on page canvas
	page->GetCanvas()->DrawImage(image, x, y, width, height);
	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}










