﻿#include "../pch.h"

using namespace std;
using namespace Spire::Pdf;

int main() {
	wstring fn1 = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ConvertImageStreamToPDF.pdf";

	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	intrusive_ptr<PdfSection> section = pdf->GetSections()->Add();
	intrusive_ptr<PdfPageBase> page = section->GetPages()->Add();

	ifstream is((fn1 + L"bg.png").c_str(), ifstream::in | ios::binary);

	is.seekg(0, is.end);
	int length = is.tellg();
	is.seekg(0, is.beg);

	char* buffer = new  char[length];
	
	is.read(buffer, length);

	intrusive_ptr<Stream> stream = new Spire::Common::Stream((unsigned char*)buffer, length);

	intrusive_ptr<PdfImage> image = new PdfImage();
	image = image->FromStream(stream);

	//Set image display location and size in PDF
	//Calculate rate
	float widthFitRate = image->GetPhysicalDimension()->GetWidth() / page->GetCanvas()->GetClientSize()->GetWidth();
	float heightFitRate = image->GetPhysicalDimension()->GetHeight() / page->GetCanvas()->GetClientSize()->GetHeight();
	float fitRate = max(widthFitRate, heightFitRate);
	//Calculate the size of image
	float fitWidth = image->GetPhysicalDimension()->GetWidth() / fitRate;
	float fitHeight = image->GetPhysicalDimension()->GetHeight() / fitRate;
	//Draw image
	page->GetCanvas()->DrawImage(image, 0, 30, fitWidth, fitHeight);
	//Save the document
	pdf->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	pdf->Close();
}


