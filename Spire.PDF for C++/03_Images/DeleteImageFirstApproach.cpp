#include "../pch.h"

using namespace std;
using namespace Spire::Pdf;

int main() {
	wstring fn1 = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DeleteImageFirstApproach.pdf";
	
	//Open pdf document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile((fn1 + L"DeleteImage.pdf").c_str());
	//Get the first page
	intrusive_ptr<PdfPageBase> page = pdf->GetPages()->GetItem(0);

	//Delete the first image on the page
	//page->DeleteImage(0);

	//Get images of the first page.
	std::vector<intrusive_ptr<PdfImageInfo>> imageInfo = page->GetImagesInfo();

	//first method
	page->DeleteImage(imageInfo[0]->GetIndex());
	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}







