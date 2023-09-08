#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"ToPDFA.pdf";
	wstring inputFile = DATAPATH"ToPDFA.pdf";

	intrusive_ptr<PdfStandardsConverter> converter_1 = new PdfStandardsConverter(inputFile.c_str());
	converter_1->ToPdfA1B(outputFile.c_str());
}
