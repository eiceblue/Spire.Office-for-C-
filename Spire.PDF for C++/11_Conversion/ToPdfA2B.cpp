#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"ToPdfA2B.pdf";
	wstring inputFile = DATAPATH"ToPdfA2B.pdf";

	intrusive_ptr<PdfStandardsConverter> converter_1 = new PdfStandardsConverter(inputFile.c_str());
	converter_1->ToPdfA2B(outputFile.c_str());
}
