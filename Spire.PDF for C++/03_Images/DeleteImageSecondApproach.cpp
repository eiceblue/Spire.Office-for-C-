#include "../pch.h"

using namespace std;
using namespace Spire::Pdf;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DeleteImage.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DeleteImageSecondApproach.pdf";
	

	//Create a pdf document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	//Load the document from disk
	pdf->LoadFromFile(inputFile.c_str());
	//Get the first page
	intrusive_ptr<PdfPageBase> page = pdf->GetPages()->GetItem(0);
	//Delete the first image on the page  second method
	page->DeleteImage(0);
	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}







