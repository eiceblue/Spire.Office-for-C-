#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"RearrangePageOrder.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	wstring inputFile = DATAPATH"SampleB_3.pdf";
	doc->LoadFromFile(inputFile.c_str());

	//Rearrange the page order
	doc->GetPages()->ReArrange(std::vector<int> {1, 0});

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}